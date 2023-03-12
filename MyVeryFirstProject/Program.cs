
// Строка 2000 - смена количества аплинковых портов на балансере

using System;
using Visio = Microsoft.Office.Interop.Visio;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Threading;
//using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Threading.Tasks;
//using Microsoft.Data.SqlClient;
using MySql.Data.MySqlClient;

namespace VeryFirstProject
{

    class Program
    {
        static int CalculateDevicesQuantity(int intLinkCounter, int intDivider)
        {
            if (intLinkCounter % intDivider == 0) return intLinkCounter / intDivider;
            else return intLinkCounter / intDivider + 1;
        }

        static int LastUsedId(MySqlCommand command, string strTableName)
        {
            object objLastUsedId = null;
            command.CommandText = $"select id from {strTableName} ORDER BY id DESC LIMIT 1;";
            command.ExecuteNonQuery();
            using (MySqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    objLastUsedId = reader.GetValue(0);
                };
            };
            return Convert.ToInt32(objLastUsedId);
        }


        static int RackObjectId(MySqlCommand command, int intRowId, string strObjectIndex, string strRackName)
        {
            command.CommandText = $"insert into Object (name, label, objtype_id) values ('{strObjectIndex} {strRackName}', {strObjectIndex}, 1560);";
            command.ExecuteNonQuery();
            int intRackId = LastUsedId(command, "Object");
            command.CommandText = $"insert into EntityLink (parent_entity_type, parent_entity_id, child_entity_type, child_entity_id) values ('row', {intRowId}, 'rack', {intRackId});";
            command.ExecuteNonQuery();
            command.CommandText = $"insert into AttributeValue (object_id, object_tid, attr_id, uint_value) values ({intRackId}, 1560, 27, 42);";  //42 вместо 96
            command.ExecuteNonQuery();
            return intRackId;
        }



        static int FillRackSlot(MySqlCommand command, string strObjectIndex, int intRackId, int intSlotNumber, int intDeviceTypeId, string strDeviceHostname)
        {

            int intMoleculeId = LastUsedId(command, "Molecule") + 1;
            command.CommandText = $"insert into Molecule (id) values ({intMoleculeId});";
            command.ExecuteNonQuery();
            command.CommandText = $"insert into Object(name, label, objtype_id) values('{strDeviceHostname}', {strObjectIndex}, {intDeviceTypeId});";
            command.ExecuteNonQuery();
            int intDeviceId = LastUsedId(command, "Object");
            command.CommandText = $"insert into Atom (molecule_id, rack_id, unit_no, atom) values ({intMoleculeId}, {intRackId}, {intSlotNumber}, 'front');";
            command.ExecuteNonQuery();
            command.CommandText = $"insert into Atom (molecule_id, rack_id, unit_no, atom) values ({intMoleculeId}, {intRackId}, {intSlotNumber}, 'interior');";
            command.ExecuteNonQuery();
            command.CommandText = $"insert into Atom (molecule_id, rack_id, unit_no, atom) values ({intMoleculeId}, {intRackId}, {intSlotNumber}, 'rear');";
            command.ExecuteNonQuery();
            ////////////////////////////////////
            command.CommandText = $"insert into RackSpace (rack_id, unit_no, atom, object_id, state) values ({intRackId}, {intSlotNumber}, 'front', {intDeviceId}, 'T');";
            command.ExecuteNonQuery();
            command.CommandText = $"insert into RackSpace (rack_id, unit_no, atom, object_id, state) values ({intRackId}, {intSlotNumber}, 'interior', {intDeviceId}, 'T');";
            command.ExecuteNonQuery();
            command.CommandText = $"insert into RackSpace (rack_id, unit_no, atom, object_id, state) values ({intRackId}, {intSlotNumber}, 'rear', {intDeviceId}, 'T');";
            command.ExecuteNonQuery();
            return intDeviceId;
        }

        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        static void Main(string[] args)
        {
            // Указываем максимальную полосу на один фильтр.
            const double doubBandwidthOnFilter4160 = 60;                                            // 4160 Divider = 60 (New Constant)     

            //~~~~~~~~~~~~~~~~~~~   Start Try Block ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            try
            {
                string connectionString;
                string strDatabaseName;

                // Подключаемся к MySQL
                Console.WriteLine("К какой БД подключаемся? 1 - тестовая; 2 - продуктивная.");
                string strDbChoice = Console.ReadLine();
                if (strDbChoice == "2")
                {
                    connectionString = "Server=192.168.50.11; Database=racktables; User Id=viktor; Password=viktor"; //192.168.50.11
                    strDatabaseName = "use racktables;";
                }
                else
                {
                    connectionString = "Server=192.168.105.250; Database=racktables_db; User Id=testuser; Password=testpassword";
                    strDatabaseName = "use racktables_db;";
                };

                Console.WriteLine("Открываем подключение.");
                MySqlCommand command = new MySqlCommand();
                MySqlConnection connection = new MySqlConnection(connectionString);
                connection.Open();
                Console.WriteLine("Подключение открыто.");
                command.Connection = connection;

                // Считываем имя файла .xlsx с вводными через интерактивный процесс. Процедура выполняется в потоке.
                string strXlsxFilePath = "";
                Thread thread1 = new Thread((ThreadStart)(() =>
                {
                    OpenFileDialog newOpenFileDialog1 = new OpenFileDialog();
                    newOpenFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
                    newOpenFileDialog1.FilterIndex = 2;
                    newOpenFileDialog1.RestoreDirectory = true;

                    if (newOpenFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        strXlsxFilePath = newOpenFileDialog1.FileName;
                    }
                }));
                thread1.SetApartmentState(ApartmentState.STA);
                thread1.Start();
                thread1.Join();

                // Из названия открытого файла формируем названия двух выводных документов.
                string strVsdFilePath = strXlsxFilePath.Replace(".xlsx", "-Layout.vsdx");
                string strExcelCablesFilePath = strXlsxFilePath.Replace(".xlsx", "-Cables.xlsx");

                // Открываем файл со считанным именем.
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(strXlsxFilePath);
                Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                Excel.Range xlRange = xlWorksheet.UsedRange;

                // Подсчиываем, сколько строк во вводном файле занято. Количество линков = количество строк - 
                int intTotalRows = xlRange.Rows.Count;
                Console.WriteLine($"Всего линков: {intTotalRows - 3}");                                                    //~~~~~~~~~~    Console Writeline (Total Links Number)

                // Из имени файла вычленяем число - ID площадки.
                string strObjectIndex = strXlsxFilePath.Substring(strXlsxFilePath.LastIndexOf('\\') + 1, strXlsxFilePath.Length - strXlsxFilePath.LastIndexOf('\\') - 1).Trim().Replace(".xlsx", "");


                // Ищем файл с IP-планом.
                //\\share\Проливка\РТ
                //string[] listFileEntries = Directory.GetFiles("\\\\share\\Проливка\\IP план");
                string[] listFileEntries = Directory.GetFiles("\\\\share\\Проливка\\РТ");
                string strFileIpPlan = "";
                string strCellText = "";

                foreach (string fileEntry in listFileEntries)
                {
                    //if (fileEntry.IndexOf("Суммарный") > 0 && fileEntry.IndexOf("~$") == -1) strFileIpPlan = fileEntry;
                    if (fileEntry.IndexOf("RT") > 0 && fileEntry.IndexOf("~$") == -1) strFileIpPlan = fileEntry;
                };

                Console.WriteLine($"Открываем IP-план: {strFileIpPlan}");
                Console.WriteLine($"Последнее изменение: {File.GetLastWriteTime(strFileIpPlan)}");

                // Открываем файл с IP-планом.
                Excel.Application xlApp3 = new Excel.Application();
                Excel.Workbook xlWorkbook3 = xlApp3.Workbooks.Open(strFileIpPlan);
                Excel.Worksheet xlWorksheet3 = (Excel.Worksheet)xlWorkbook3.Worksheets.get_Item(1);

                //Проход по строкам. При обнаружении совпадения номера площадки считываем в переменную расширенный индекс площадки.
                Console.WriteLine($"Строк в IP-плане: {xlWorksheet3.UsedRange.Rows.Count}");
                string strFullSiteIndex = null;
                string strOperatorName = "";
                string strObjectAddress = "";
                if (strObjectIndex[0] == '0') strObjectIndex = strObjectIndex.Substring(1);
                if (strObjectIndex[0] == '0') strObjectIndex = strObjectIndex.Substring(1);
                for (int intCurrentRow = 2; intCurrentRow <= xlWorksheet3.UsedRange.Rows.Count; intCurrentRow++)
                {
                    strCellText = ((Excel.Range)xlWorksheet3.Cells[intCurrentRow, 1]).Value2.ToString();
                    if (strCellText == strObjectIndex)
                    {
                        strFullSiteIndex = ((Excel.Range)xlWorksheet3.Cells[intCurrentRow, 5]).Value2.ToString();
                        strOperatorName = ((Excel.Range)xlWorksheet3.Cells[intCurrentRow, 2]).Value2.ToString();
                        strObjectAddress = ((Excel.Range)xlWorksheet3.Cells[intCurrentRow, 3]).Value2.ToString() + ", " + ((Excel.Range)xlWorksheet3.Cells[intCurrentRow, 4]).Value2.ToString();
                        Console.WriteLine($"Индекс площадки: {strFullSiteIndex}");
                        break;
                    }
                };
                if (strFullSiteIndex == null && strDatabaseName == "use racktables;") throw new Exception("Площадка отсутствует в RackTables!");

                // Закрываем эксель с IP-планом
                xlWorkbook3.Close();
                xlApp3.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkbook3);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp3);


                // Создаём списки словарей LAN-WAN
                List<Dictionary<string, object>> listLanDevices = new List<Dictionary<string, object>>();           //LAN Devices
                List<Dictionary<string, object>> listWanDevices = new List<Dictionary<string, object>>();           //WAN Devices
                List<Dictionary<string, object>> listLanPorts = new List<Dictionary<string, object>>();             //LAN Ports
                List<Dictionary<string, object>> listWanPorts = new List<Dictionary<string, object>>();             //WAN Ports

                // Создаём списки словарей для КЖ
                List<Dictionary<string, string>> listCableJournal_1 = new List<Dictionary<string, string>>();                    //Router-Bypass
                List<Dictionary<string, string>> list_CableJournal_Bypass_Filter = new List<Dictionary<string, string>>();       //Bypass-Balancer
                List<Dictionary<string, string>> listCableJournal_3 = new List<Dictionary<string, string>>();                    //Balancer-Filter
                List<Dictionary<string, string>> listCableJournal_4 = new List<Dictionary<string, string>>();                    //Для ТП
                List<Dictionary<string, string>> listCableJournal_Management = new List<Dictionary<string, string>>();           //Management
                List<Dictionary<string, string>> listCableJournal_Log = new List<Dictionary<string, string>>();                  //Log
                List<Dictionary<string, string>> listCableJournal_Interconnect = new List<Dictionary<string, string>>();         //Interconnect
                List<Dictionary<string, string>> list_CableJournal_LAN_Bypass = new List<Dictionary<string, string>>();          //Общий (LAN-Bypass)
                List<Dictionary<string, string>> list_CableJournal_WAN_Bypass = new List<Dictionary<string, string>>();          //Общий (WAN-Bypass)
                List<Dictionary<string, string>> list_CableJournal_Bypass_Balancer = new List<Dictionary<string, string>>();     //list_CableJournal_Bypass_Balancer
                List<Dictionary<string, string>> list_CableJournal_Balancer_Filter = new List<Dictionary<string, string>>();     //list_CableJournal_Balancer_Filter
                List<Dictionary<string, string>> list_CableJournal_Highway_Peremychka = new List<Dictionary<string, string>>();  //list_CableJournal_Highway_Peremychka

                // Создаём массивы словарей для КЖ
                Dictionary<string, string>[,] arrCableJournal_LAN_Bypass = new Dictionary<string, string>[200, 200];
                Dictionary<string, string>[,] arrCableJournal_WAN_Bypass = new Dictionary<string, string>[200, 200];
                Dictionary<string, string>[,,] arr_CableJournal_Bypass_Balancer = new Dictionary<string, string>[100, 200, 10];          //Bypass-Balancer
                Dictionary<string, string>[,,] arrCableJournal_Balancer_Filter = new Dictionary<string, string>[100, 200, 20];           //Balancer-Filter
                Dictionary<string, string>[] arrCableJournal_Bypass_Filter = new Dictionary<string, string>[200];                        //Bypass-Filter
                Dictionary<string, string>[,] arrCableJournal_Highway_Peremychka = new Dictionary<string, string>[100, 100];             
                //Balancer Peremychka

                // Создаём список для спецификации (Удалить!)
                List<string> list_Specification = new List<string>();

                // Создаём списки фигур Visio
                List<Visio.Shape> listShapesLanDevices = new List<Visio.Shape>();                                    //LAN Devices Rects       
                List<Visio.Shape> listShapesWanDevices = new List<Visio.Shape>();                                    //WAN Devices Rects               
                List<Visio.Shape> listShapesLanPorts = new List<Visio.Shape>();                                      //LAN Ports Rects 
                List<Visio.Shape> listShapesWanPorts = new List<Visio.Shape>();                                      //LAN Ports Rects 

                // Создаём массивы фигур Visio
                Visio.Shape[] arrShapesLanDevices = new Visio.Shape[20];
                Visio.Shape[] arrShapesWanDevices = new Visio.Shape[20];
                Visio.Shape[] arrShapesLanPorts = new Visio.Shape[200];
                Visio.Shape[] arrShapesWanPorts = new Visio.Shape[200];
                Visio.Shape[] arrShapesLanFakeLines = new Visio.Shape[200];
                Visio.Shape[] arrShapesWanFakeLines = new Visio.Shape[200];
                Visio.Shape[] arrShapesLanFakeCircles = new Visio.Shape[200];
                Visio.Shape[] arrShapesWanFakeCircles = new Visio.Shape[200];
                Visio.Shape[] arrShapesLanFakeConnections = new Visio.Shape[200];
                Visio.Shape[] arrShapesWanFakeConnections = new Visio.Shape[200];


                // Создаём массив аплинк-портов балансера (почему здесь?)
                int[] arrUplinkPortsOnBalancer = new int[200];

                // Считываем информацию о площадке из верхней строки.
                double doubSummaryBandwidth = Convert.ToDouble(((Excel.Range)xlWorksheet.Cells[2, 19]).Value2.ToString());         // Read Summary BW from File [2,7]
                string strBypassModel = ((Excel.Range)xlWorksheet.Cells[2, 20]).Value2.ToString();                                 // Read Bypass Model     (Не было)
                string strFilterModel = ((Excel.Range)xlWorksheet.Cells[2, 21]).Value2.ToString();                                 // Read Filter Model         [2,19]
                string strContinentModel = ((Excel.Range)xlWorksheet.Cells[2, 22]).Value2.ToString();                              // Read Continent Model  (Не было)
                //string strObjectAddress = ((Excel.Range)xlWorksheet.Cells[2, 16]).Value2.ToString();                             // Read Object Post Address  [2,16]
                string strAssimetry = ((Excel.Range)xlWorksheet.Cells[2, 23]).Value2.ToString();                                   // Read Assimetry (или отметка об эшелоне)
                string strBalancerNumberFromInput = ((Excel.Range)xlWorksheet.Cells[2, 24]).Value2.ToString();                     // Read Balancers Number (не было)
                string strFilterNumberFromInput = ((Excel.Range)xlWorksheet.Cells[2, 25]).Value2.ToString();                       // Read Filter Number (не было)

                // Создаём переменные для балансировщика (перенести ниже!)
                int intCurrentPortInChassis;
                int intStartBalancerPort;
                int intTotalBalancers = Convert.ToInt32(strBalancerNumberFromInput);

                // Создаём глобальные логические переменные (перенести ниже!)
                bool boolEolBypass = false;
                bool boolContinentIpcR300 = false;
                bool boolEshelon = false;

                // Переписывание голобальных логических переменных: 
                if (strBypassModel == "IBS1UP") boolEolBypass = true;
                if (strContinentModel == "300" || strContinentModel == "550") boolContinentIpcR300 = true;
                if (strAssimetry == "эшелон" || strAssimetry == "эшкрест") boolEshelon = true;

                // Добавляем в БД площадку (location).
                command.CommandText = strDatabaseName;
                command.ExecuteNonQuery();

                //Проверяем в базе наличие название провайдера в Location
                int intOccupiedRows = 0;
                command.CommandText = $"SELECT COUNT(*) from Object where name = '{strOperatorName}';";
                using (MySqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        intOccupiedRows = Convert.ToInt32(reader.GetValue(0));
                    };
                };
                //Console.WriteLine($"Оператор {strOperatorName} присутствует {intOccupiedRaws} раз.");

                int intLocationId = 0;
                if (intOccupiedRows > 0)
                {
                    command.CommandText = $"SELECT id from Object where objtype_id = '1562' AND name = '{strOperatorName}';";
                    using (MySqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            intLocationId = Convert.ToInt32(reader.GetValue(0));
                        };
                    };
                    Console.WriteLine($"Оператор {strOperatorName} уже присутствует в системе.");
                }
                else
                {
                    //Добавляем оператора.
                    command.CommandText = $"insert into Object (name, objtype_id) values ('{strOperatorName}', 1562);";
                    command.ExecuteNonQuery();
                    // Чтение Location ID. Расчёт Row ID. Обращение к функции.
                    intLocationId = LastUsedId(command, "Object");
                    Console.WriteLine($"Добавляем оператор {strOperatorName} в систему.");
                };


                // Добавление ряда в площадку. Создание линка между записями.
                command.CommandText = $"insert into Object (name, label, objtype_id) values ('{strObjectIndex} {strObjectAddress}, {DateTime.Now:g} scriptgen', {strObjectIndex}, 1561);";   //{DateTime.Now.ToLongDateString()
                command.ExecuteNonQuery();
                int intRowId = LastUsedId(command, "Object");

                command.CommandText = $"insert into EntityLink (parent_entity_type, parent_entity_id, child_entity_type, child_entity_id) values ('location', {intLocationId}, 'row', {intRowId});";
                command.ExecuteNonQuery();

                // Добавление в созданный выше ряд нескольких стоек.
                int intBypassRackId = RackObjectId(command, intRowId, strObjectIndex, "Байпасы");
                int intBalancerRackId = RackObjectId(command, intRowId, strObjectIndex, "Балансеры");
                int intFilterRackId = RackObjectId(command, intRowId, strObjectIndex, "Фильтры");
                int intNetSrvRackId = RackObjectId(command, intRowId, strObjectIndex, "Менеджмент");
                int intLanWanRackId = RackObjectId(command, intRowId, strObjectIndex, "LAN-WAN");

                // Расчёт количества фильтров и СПФС из суммарного трафика площадки.
                int intTotalFiltersFromBw = Convert.ToInt32(Math.Ceiling(doubSummaryBandwidth / doubBandwidthOnFilter4160));        //Calculated Total Filters Number (4120 or 4160)
                int intTotalLogServers = Convert.ToInt32(Math.Ceiling(doubSummaryBandwidth / 350));                                 //Calculated Total SPHD Number
                Console.WriteLine($"Количество СПФС: {intTotalLogServers}");

                // Расчёт количества гидр на фильтре
                int intHydrasOnFilter = 4;                                      //Default = 4160
                if (strFilterModel == "4120") intHydrasOnFilter = 3;            // 4120

                // Расчёт количества балансировщиков исходя из количества фильтров (Разобраться!!!)
                int intBalancersFromFiltersAndHydras = CalculateDevicesQuantity(intTotalFiltersFromBw * intHydrasOnFilter, 16);
                int intBalancersFromBw = CalculateDevicesQuantity(intTotalFiltersFromBw, intHydrasOnFilter);
                if (intTotalFiltersFromBw == 1) intBalancersFromBw = 0;

                // Определение количества портов на одиночном фильтре исходя из модели фильтра.
                int intPortsNumberOnSingleFilter = 0;
                switch (strFilterModel)
                {
                    case "4160":
                        intPortsNumberOnSingleFilter = 16;
                        break;
                    case "4120":
                        intPortsNumberOnSingleFilter = 12;
                        break;
                    case "4080":
                        intPortsNumberOnSingleFilter = 8;
                        break;
                };

                // Определение кучи переменных разных типов
                bool boolLanObjectNotFound;
                bool boolWanObjectNotFound;
                int intCurrentJournalItem = 0;
                int intLocalPortCounter = 0;
                int intCurrentLanPortsCounter;
                int intCurrentWanPortsCounter;
                int intCurrent100mBalancerPort;
                int intTotal100mBalancerPorts;
                int intCurrentGlobalDeviceIndex = 0;
                int intCurrentGlobalPortIndex = 0;
                int intCurrentGlobalDeviceIndexForPort = 0;
                int intGlobalCableCounter = 0;
                int intLinkCounter100 = 0;
                int intLinkCounter40 = 0;
                int intLinkCounter10 = 0;
                int intLinkCounter1Fiber = 0;
                int intLinkCounter1Copper = 0;
                int intLinkCounterEol = 0;
                int intLinkCounter10Old = 0;
                int intLinkCounter10New = 0;
                int intCurrentOverallLinkNumber = 0;
                int intCurrentLagNumber = 0;
                int[] arrLanLagCounter = new int[200];
                int[] arrLagIteration = new int[100];
                string strCableInHydra;
                string strCurrentLanHostname;
                string strCurrentWanHostname;
                string strCurrentLanPortName;
                string strCurrentWanPortName;
                string strCurrentLinkType;
                string strLanOdfLocation;
                string strLanOdfPort;
                string strWanOdfLocation;
                string strWanOdfPort;
                string strCurrentLinkStatus;
                string strLanLagName;
                string strPreviousHostname = "";
                string strPreviousLag = "";
                string LanWanComment = "";
                string SideADevice = "";

                // Проход по всем строкам исходного .xlsx, описывающих линки.
                for (int intCurrentRow = 4; intCurrentRow <= intTotalRows; intCurrentRow++)
                {
                    boolLanObjectNotFound = true;
                    boolWanObjectNotFound = true;
                    strCurrentLinkStatus = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 2]).Value2.ToString();
                    strCurrentLinkType = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 3]).Value2.ToString();
                    strCurrentLanHostname = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 4]).Value2.ToString();
                    strCurrentLanPortName = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 5]).Value2.ToString();
                    strCurrentWanHostname = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 8]).Value2.ToString();
                    strCurrentWanPortName = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 9]).Value2.ToString();
                    strLanOdfLocation = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 12]).Value2.ToString().Trim();
                    strLanOdfPort = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 13]).Value2.ToString();
                    strWanOdfLocation = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 14]).Value2.ToString();
                    strWanOdfPort = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 15]).Value2.ToString();
                    strLanLagName = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 7]).Value2.ToString();
                    intCurrentOverallLinkNumber++;

                    // Подсчёт количества используемых линков каждого типа.
                    switch (strCurrentLinkType)
                    {
                        case "100G":
                            intLinkCounter100++;
                            break;
                        case "40G":
                            intLinkCounter40++;
                            break;
                        case "10G":
                            intLinkCounter10++;
                            break;
                        case "1G":
                            intLinkCounter1Fiber++;
                            break;
                        case "1Copper":
                            intLinkCounter1Copper++;
                            break;
                        case "10_EOL":
                            intLinkCounterEol++;            // Потом удалить!
                            break;
                    }

                    // Подсчёт количества используемых старых и новых линков 10G. Для подсчёта количества байпасов разных моделей.
                    switch (strCurrentLinkStatus)
                    {
                        case "через ТСПУ":
                            intLinkCounter10Old++;
                            break;
                        case "план":
                            intLinkCounter10New++;
                            break;
                    }

                    // Устройства LAN.
                    // Если имя устройства найдено, добавляем порт в текущий словарь LAN-устройств.
                    if (listLanDevices.Count > 0)
                    {
                        foreach (Dictionary<string, object> dictLanDevices in listLanDevices)
                        {
                            if (dictLanDevices.ContainsValue(strCurrentLanHostname))
                            {
                                intCurrentLanPortsCounter = Convert.ToInt32(dictLanDevices["Ports_Number"]);
                                intCurrentLanPortsCounter++;
                                dictLanDevices["Ports_Number"] = intCurrentLanPortsCounter;
                                boolLanObjectNotFound = false;
                                intCurrentGlobalDeviceIndexForPort = Convert.ToInt32(dictLanDevices["Device_Index"]);
                                break;
                            };
                        };
                    };

                    // Если имя устройства не найдено, добавляем устройство и порт в текущий словарь этого LAN-устройства под номером 1.
                    if (boolLanObjectNotFound)
                    {
                        listLanDevices.Add(new Dictionary<string, object>());
                        intCurrentGlobalDeviceIndex++;
                        listLanDevices[listLanDevices.Count - 1].Add("Device_Name", strCurrentLanHostname);
                        listLanDevices[listLanDevices.Count - 1].Add("Ports_Number", 1);
                        listLanDevices[listLanDevices.Count - 1].Add("Device_Index", intCurrentGlobalDeviceIndex);
                        intCurrentGlobalDeviceIndexForPort = intCurrentGlobalDeviceIndex;
                    };

                    // Подсчёт количества используемых портов LAN, разных типов.
                    // Заполнение словаря портов устройств LAN.
                    listLanPorts.Add(new Dictionary<string, object>());
                    intCurrentGlobalPortIndex++;
                    switch (strCurrentLinkType)
                    {
                        case "100G":
                            intLocalPortCounter = 2 * intLinkCounter100 - 1;
                            break;
                        case "40G":
                            intLocalPortCounter = 2 * intLinkCounter40 - 1;
                            break;
                        case "10G":
                            intLocalPortCounter = 2 * intLinkCounter10 - 1;
                            break;
                        case "1G":
                            intLocalPortCounter = 2 * intLinkCounter1Fiber - 1;
                            break;
                        case "1Copper":
                            intLocalPortCounter = 2 * intLinkCounter1Copper - 1;
                            break;
                        case "10_EOL":
                            intLocalPortCounter = 2 * intLinkCounterEol - 1;                                                //Убрать после того как переделаю логику
                            break;
                    }
                    listLanPorts[listLanPorts.Count - 1].Add("Device_Index", intCurrentGlobalDeviceIndexForPort);
                    listLanPorts[listLanPorts.Count - 1].Add("Port_Name", strCurrentLanPortName);
                    listLanPorts[listLanPorts.Count - 1].Add("Port_Index", intLocalPortCounter);                          //~~~~~~~~~~~~~~~~~~    Поменять!   ~~~~~~~~~~~~~~
                    listLanPorts[listLanPorts.Count - 1].Add("Link_Type", strCurrentLinkType);
                    listLanPorts[listLanPorts.Count - 1].Add("Overall_Link_Number", intCurrentOverallLinkNumber);

                    // Запись LAN-портов в массив для КЖ LAN-Bypass.
                    LanWanComment = strCurrentLanHostname + "\n" + Convert.ToString(strCurrentLanPortName);
                    if (strLanOdfLocation == "без ODF" || strLanOdfLocation == "Без ODF" || strLanOdfLocation == "") 
                        { strLanOdfLocation = strCurrentLanHostname; strLanOdfPort = strCurrentLanPortName; LanWanComment = "Прямое включение"; };
                    intCurrentJournalItem++;
                    intGlobalCableCounter++;
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0] = new Dictionary<string, string>();
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Port_ID", Convert.ToString(intLocalPortCounter));
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_A_Name", strLanOdfLocation);
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_A_Port", strLanOdfPort);
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Comment", LanWanComment);
                    //arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Comment", strCurrentLanHostname + "\n" + Convert.ToString(strCurrentLanPortName));


                    // Устройства WAN.
                    // Если имя устройства найдено, добавляем порт в тукущий словарь WAN-устройств.
                    if (listWanDevices.Count > 0)
                    {
                        foreach (Dictionary<string, object> dictWanDevices in listWanDevices)
                        {
                            if (dictWanDevices.ContainsValue(strCurrentWanHostname))
                            {
                                intCurrentWanPortsCounter = Convert.ToInt32(dictWanDevices["Ports_Number"]);
                                intCurrentWanPortsCounter++;
                                dictWanDevices["Ports_Number"] = intCurrentWanPortsCounter;
                                boolWanObjectNotFound = false;
                                intCurrentGlobalDeviceIndexForPort = Convert.ToInt32(dictWanDevices["Device_Index"]);
                                break;
                            };
                        };
                    };

                    // Если имя устройства не найдено, добавляем устройство и порт в текущий словарь этого LAN-устройства под номером 1.
                    if (boolWanObjectNotFound)
                    {
                        listWanDevices.Add(new Dictionary<string, object>());
                        intCurrentGlobalDeviceIndex++;
                        listWanDevices[listWanDevices.Count - 1].Add("Device_Name", strCurrentWanHostname);
                        listWanDevices[listWanDevices.Count - 1].Add("Ports_Number", 1);
                        listWanDevices[listWanDevices.Count - 1].Add("Device_Index", intCurrentGlobalDeviceIndex);
                        intCurrentGlobalDeviceIndexForPort = intCurrentGlobalDeviceIndex;

                    };

                    // Подсчёт количества используемых портов WAN, разных типов.
                    // Заполнение словаря портов устройств WAN.
                    listWanPorts.Add(new Dictionary<string, object>());
                    intCurrentGlobalPortIndex++;

                    switch (strCurrentLinkType)
                    {
                        case "100G":
                            intLocalPortCounter = 2 * intLinkCounter100;
                            break;
                        case "40G":
                            intLocalPortCounter = 2 * intLinkCounter40;
                            break;
                        case "10G":
                            intLocalPortCounter = 2 * intLinkCounter10;
                            break;
                        case "1G":
                            intLocalPortCounter = 2 * intLinkCounter1Fiber;
                            break;
                        case "1Copper":
                            intLocalPortCounter = 2 * intLinkCounter1Copper;
                            break;
                        case "10_EOL":
                            intLocalPortCounter = 2 * intLinkCounterEol;
                            break;
                    }

                    // Запись WAN-портов в массив для КЖ WAN-Bypass.
                    listWanPorts[listWanPorts.Count - 1].Add("Device_Index", intCurrentGlobalDeviceIndexForPort);
                    listWanPorts[listWanPorts.Count - 1].Add("Port_Name", strCurrentWanPortName);
                    listWanPorts[listWanPorts.Count - 1].Add("Port_Index", intLocalPortCounter);          //~~~~~~~~~~~~~~~~~~    Поменять!   ~~~~~~~~~~~~~~
                    listWanPorts[listWanPorts.Count - 1].Add("Link_Type", strCurrentLinkType);
                    listWanPorts[listWanPorts.Count - 1].Add("Overall_Link_Number", intCurrentOverallLinkNumber);

                    // Запись WAN-портов в массив для КЖ WAN-Bypass.
                    LanWanComment = strCurrentWanHostname + "\n" + Convert.ToString(strCurrentWanPortName);
                    if (strWanOdfLocation == "без ODF" || strWanOdfLocation == "Без ODF" || strWanOdfLocation == "") { strWanOdfLocation = strCurrentWanHostname; strWanOdfPort = strCurrentWanPortName; LanWanComment = "Прямое включение"; };
                    intCurrentJournalItem++;
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1] = new Dictionary<string, string>();
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Port_ID", Convert.ToString(intLocalPortCounter));
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_A_Name", strWanOdfLocation);
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_A_Port", strWanOdfPort);
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Comment", LanWanComment);
                    //arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Comment", strCurrentWanHostname + "\n" + Convert.ToString(strCurrentWanPortName));


                    // Работа с LAG. Ориентируемся на то, что лаги друг под другом будут.
                    // Если имя LAG совпадает с использованным на прошлой итерации, увеличиваем счётчик линков для текущего LAG.
                    // Иначе двигаем счётчик лагов и также увеличиваем счётчки линков.
                    //if (strLanLagName != strPreviousLag)
                    if (strCurrentLanHostname != strPreviousHostname || strLanLagName != strPreviousLag)
                    {
                        intCurrentLagNumber++;
                    };

                    arrLanLagCounter[intCurrentLagNumber]++;
                    strPreviousHostname = strCurrentLanHostname;
                    strPreviousLag = strLanLagName;
                    // Конец прохода по строке линков в эксель-файле.
                };

                // Закрываем эксель с шаблоном
                xlWorkbook.Close();
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);

                // Для смешанного состава байсов считаем отдельно количество IS40 и IBS1UP
                int intBypasses10Old = CalculateDevicesQuantity(intLinkCounter10Old, 4);
                int intBypasses10New = CalculateDevicesQuantity(intLinkCounter10New, 6);

                // Вычисление общего количества линков.
                int intTotalOverallLinkNumber = intTotalRows - 3;

                // Вывод на консоль числа линков в каждом из лагов.
                for (int intLagCounter = 1; intLagCounter <= intCurrentLagNumber; intLagCounter++)
                {
                    Console.WriteLine($"LAG #{intLagCounter}, линков {arrLanLagCounter[intLagCounter]}.");
                };

                // Вывод на консоль числа линков каждого типа.
                if (intLinkCounter100 > 0) Console.WriteLine($"Порты 100G: {intLinkCounter100}");
                if (intLinkCounter40 > 0) Console.WriteLine($"Порты 40G: {intLinkCounter40}");
                if (intLinkCounter10 > 0) Console.WriteLine($"Порты 10G: {intLinkCounter10}");
                if (intLinkCounter1Fiber > 0) Console.WriteLine($"Порты 1G (оптика): {intLinkCounter1Fiber}");
                if (intLinkCounter1Copper > 0) Console.WriteLine($"Порты 1G (медь): {intLinkCounter1Copper}");

                // Вывод на консоль числа типа включения (прямое/крестовое).
                bool boolCrossLayout;
                if (intLinkCounter100 == 0 && intLinkCounter40 == 0)
                {
                    boolCrossLayout = false;
                    Console.WriteLine("Прямая схема.");
                }
                else
                {
                    boolCrossLayout = true;
                    Console.WriteLine("Крестовая схема.");
                };

                // Определение количества байпасов каждого типа.
                // Делим количество линков нужного типа на делитель, соответствующий количеству линков на шасси.
                int intBypassPortDivider;
                if (boolEolBypass) intBypassPortDivider = 4;
                else intBypassPortDivider = 6;
                int intTotalIs100Bypasses = CalculateDevicesQuantity(intLinkCounter100, 2);
                int intTotalSignaltek100Bypasses = CalculateDevicesQuantity(intLinkCounter100, 4);
                int intTotalIs40Bypasses = CalculateDevicesQuantity(intLinkCounter40, 3);
                int intTotalIs10Bypasses = CalculateDevicesQuantity(intLinkCounter10 + intLinkCounter1Fiber, intBypassPortDivider);
                int intTotalIs1CopperBypasses = CalculateDevicesQuantity(intLinkCounter1Copper, 4);
                int intTotalIbs1upBypasses = CalculateDevicesQuantity(intLinkCounter10, 4);                                       //   Потом удалить!

                // Вывод на консоль количества сипользуемых байпасов.
                if (intTotalIs100Bypasses > 0) Console.WriteLine($"IS100: {intTotalIs100Bypasses}");
                if (intTotalIs40Bypasses > 0 | (intTotalIs10Bypasses > 0) && !boolEolBypass) Console.WriteLine($"IS40 Number: {intTotalIs40Bypasses} + {intTotalIs10Bypasses} = {intTotalIs40Bypasses + intTotalIs10Bypasses}");
                if ((intTotalIs10Bypasses > 0) && boolEolBypass) Console.WriteLine($"IBS1UP Number: {intTotalIs10Bypasses}");

                // Создание переменных координат.
                double doubStartPointNextShapeX;
                double doubStartPointNextShapeY;
                double doubDeviceStartPointX = 0;
                double doubDeviceStartPointY;
                double doubDeviceEndPointX;
                double doubDeviceEndPointY;
                double doubShiftY = 0;
                int intPortsOnDevice;

                // Определение балансерная схема или без балансерная.              Так как количество балансеров выставляется в шаблоне вручную, можно эти расчёты не использовать. (!!!)
                // Схема считается безбалансерной, если линков 10G более 0 и менее 8.
                bool boolNoBalancer = false;
                if (intLinkCounter10 > 0 & intLinkCounter10 <= 8) boolNoBalancer = true;
                if (intLinkCounter1Fiber > 0 & intLinkCounter10 == 0) boolNoBalancer = true;
                if (intTotalBalancers > 0) boolNoBalancer = false;
                if (boolNoBalancer) Console.WriteLine("Прямая коммутация байпасов к фильтру. Без балансировщиков.");

                // Определение количества балансировщиков.          Так как количество балансеров выставляется в шаблоне вручную, можно эти расчёты не использовать. (!!!)
                int intBalancersFromPorts = CalculateDevicesQuantity(2 * intLinkCounter100 + 2 * intLinkCounter40 + intLinkCounter10 / 2, 16);
                int intBalancersFinalQuantity = Math.Max(intBalancersFromBw, intBalancersFromPorts);

                //Установка начальной координаты: 1,1.
                doubStartPointNextShapeX = 1;
                doubStartPointNextShapeY = 1;

                // Создание переменной счётчика индексов устройств LAN-WAN в словарях. 
                int intListCurrentIndex = 0;

                // Добавление в массив словарей координат для портов LAN-устройств.
                foreach (Dictionary<string, object> dictLanDevices in listLanDevices)
                {
                    intPortsOnDevice = Convert.ToInt32(dictLanDevices["Ports_Number"]);
                    if (intPortsOnDevice < 3) intPortsOnDevice = 3;
                    doubDeviceStartPointX = doubStartPointNextShapeX;
                    doubDeviceStartPointY = doubStartPointNextShapeY;
                    doubDeviceEndPointX = doubDeviceStartPointX + 0.2 * intPortsOnDevice + 0.1;
                    doubDeviceEndPointY = doubDeviceStartPointY + 1;
                    doubStartPointNextShapeX = doubDeviceEndPointX + 1.5;
                    doubStartPointNextShapeY = doubDeviceStartPointY;

                    listLanDevices[intListCurrentIndex].Add("StartX", doubDeviceStartPointX);
                    listLanDevices[intListCurrentIndex].Add("StartY", doubDeviceStartPointY);
                    listLanDevices[intListCurrentIndex].Add("EndX", doubDeviceEndPointX);
                    listLanDevices[intListCurrentIndex].Add("EndY", doubDeviceEndPointY);

                    intListCurrentIndex++;
                };

                // Вычисление расстояния между устройствами LAN и WAN, координаты Y устройств WAN.
                doubStartPointNextShapeX = 1;
                doubStartPointNextShapeY = intTotalIs100Bypasses * 12 + intTotalIs40Bypasses * 8 + intTotalIs10Bypasses * 16 + intTotalIs1CopperBypasses * 6; // + intTotalIbs1upBypasses * 6;

                // Сброс индекса портов WAN-устройств
                intListCurrentIndex = 0;

                // Добавление в массив словарей координат для портов WAN-устройств.
                foreach (Dictionary<string, object> dictWanDevices in listWanDevices)
                {
                    intPortsOnDevice = Convert.ToInt32(dictWanDevices["Ports_Number"]);
                    if (intPortsOnDevice < 3) intPortsOnDevice = 3;
                    doubDeviceStartPointX = doubStartPointNextShapeX + 0.3;
                    doubDeviceStartPointY = doubStartPointNextShapeY;
                    doubDeviceEndPointX = doubDeviceStartPointX + 0.2 * intPortsOnDevice + 0.2;
                    doubDeviceEndPointY = doubDeviceStartPointY + 1;
                    doubStartPointNextShapeX = doubDeviceEndPointX + 1.5;
                    doubStartPointNextShapeY = doubDeviceStartPointY;
                    listWanDevices[intListCurrentIndex].Add("StartX", doubDeviceStartPointX);
                    listWanDevices[intListCurrentIndex].Add("StartY", doubDeviceStartPointY);
                    listWanDevices[intListCurrentIndex].Add("EndX", doubDeviceEndPointX);
                    listWanDevices[intListCurrentIndex].Add("EndY", doubDeviceEndPointY);
                    intListCurrentIndex++;
                };

                // Опускание текущей координаты Y: байпасы должны быть чуть ниже, чем устройства WAN.
                doubStartPointNextShapeY -= 2.1;

                // Создание файла визио.
                Visio.Application app = new Visio.Application();
                Visio.Document doc = app.Documents.Add("");
                Visio.Page page1 = doc.Pages[1];
                page1.Name = "Схема Кабельных Соединений";

                // Создание сущности "выбора нескольких фигур".
                Visio.Selection vsoSelection;
                Visio.Window vsoWindow;
                vsoWindow = app.ActiveWindow;

                // Создание новых слоёв Visio. Сейчас не используется. (!!!)
                Visio.Layers vsoAllLayers = page1.Layers;
                Visio.Layer vsoLayerMgmt = vsoAllLayers.Add("Management");
                Visio.Layer vsoLayerMain = vsoAllLayers.Add("Main");

                // Объявление переменных координат фигур.
                double doubNextPortStartPointX;
                double doubNextPortStartPointY;
                double doubLanLastX;
                double doubWanLastX;

                // Объявление переменных в рисовании фигур.
                string strShapeName;
                string strNameForRackTables;
                string strPortName;
                int intPortOnChassis = 0;
                int intCurrentLanChassis = 0;
                int intCurrentWanChassis = 0;
                string strLinkTypeId = "";
                string strLinkSubTypeId = "";
                string strRouterNameModified = "";

                // Переменные для работы с MySQL.
                int intCurrentRackSlot;
                int intDeviceId;

                ////////////////////////////////////////////////////////////////////////////////// Рисуем Шасси LAN ////////////////////////////////////////////////////////////////////////////

                // Сброс счётчика линков перед рисование фигур LAN Visio
                intCurrentOverallLinkNumber = 0;
                //int TotalLanDevices = listLanDevices.Count;

                // Счётчик слотов в стойке LAN-WAN выставляем в 0. Заполнение стойки LAN-устройствами начинаем с самого низа.
                intCurrentRackSlot = 10;

                // Проход по словарю LAN-устройств
                foreach (Dictionary<string, object> dictLanDevices in listLanDevices)
                {
                    intCurrentLanChassis++;
                    doubDeviceStartPointX = Convert.ToDouble(dictLanDevices["StartX"]);
                    doubDeviceStartPointY = Convert.ToDouble(dictLanDevices["StartY"]);
                    doubDeviceEndPointX = Convert.ToDouble(dictLanDevices["EndX"]);
                    doubDeviceEndPointY = Convert.ToDouble(dictLanDevices["EndY"]);
                    strShapeName = $"{Convert.ToString(dictLanDevices["Device_Name"])}";
                    strRouterNameModified = Convert.ToString(dictLanDevices["Device_Name"]).Replace(" ", "_").Replace("-", "_");
                    strNameForRackTables = $"{strFullSiteIndex}-LAN-{strRouterNameModified}";
                    listShapesLanDevices.Add(page1.DrawRectangle(doubDeviceStartPointX, doubDeviceStartPointY, doubDeviceEndPointX + 0.1, doubDeviceEndPointY));
                    listShapesLanDevices[listShapesLanDevices.Count - 1].Text = strShapeName;
                    listShapesLanDevices[listShapesLanDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(169,169,169)";
                    listShapesLanDevices[listShapesLanDevices.Count - 1].Data3 = Convert.ToString(intCurrentLanChassis);

                    // Добавляем шасси LAN в БД. Стойку LAN-WAN заплняем снизу вверх.
                    intCurrentRackSlot++;
                    intDeviceId = FillRackSlot(command, strObjectIndex, intLanWanRackId, intCurrentRackSlot, 50004, strNameForRackTables);

                    // Устанавливаем координаты первого порта.
                    doubNextPortStartPointX = doubDeviceStartPointX - 0.1;
                    doubNextPortStartPointY = doubDeviceEndPointY + 0.2;

                    // Проход по словарю портов LAN-устройств
                    foreach (Dictionary<string, object> dictLanPorts in listLanPorts)
                    {
                        if (Convert.ToInt32(dictLanDevices["Device_Index"]) == Convert.ToInt32(dictLanPorts["Device_Index"]))
                        {
                            switch (dictLanPorts["Link_Type"])
                            {
                                case "100G":
                                    strLinkTypeId = "15";
                                    strLinkSubTypeId = "1670";              // Переписать логику. Сделать отдельные switch case для субтайпов
                                    break;
                                case "40G":
                                    strLinkTypeId = "10";
                                    strLinkSubTypeId = "1664";
                                    break;
                                case "10G":                                 /////////////////////////       Type = 1670 - только для 100GBASE-LR4! Для SR надо дописать алгоритм.
                                    strLinkTypeId = "9";
                                    strLinkSubTypeId = "36";
                                    break;
                                case "1G":
                                    strLinkTypeId = "4";
                                    strLinkSubTypeId = "1204";                //Узнать!
                                    break;
                            }


                            // Рисование порта с добавлением его в массив фигур LAN. Индекс порта в массиве берётся из списка словарей LAN-портов.
                            intCurrentOverallLinkNumber = Convert.ToInt32(dictLanPorts["Overall_Link_Number"]);
                            strPortName = Convert.ToString(dictLanPorts["Port_Name"]);
                            intPortOnChassis++;
                            arrShapesLanPorts[intCurrentOverallLinkNumber] = page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.6, doubNextPortStartPointY + 0.2);
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Data2 = Convert.ToString(dictLanPorts["Link_Type"]);     // Проверить, используются ли Data2 и Data3.
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Data3 = Convert.ToString(intCurrentLanChassis);
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Text = strPortName;
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Rotate90();
                            arrShapesLanPorts[intCurrentOverallLinkNumber].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            arrShapesLanFakeLines[intCurrentOverallLinkNumber] = page1.DrawLine(doubNextPortStartPointX + 0.3, doubNextPortStartPointY + 0.4, doubNextPortStartPointX + 0.3, doubNextPortStartPointY + 3.6);
                            arrShapesLanFakeCircles[intCurrentOverallLinkNumber] = page1.DrawOval(doubNextPortStartPointX + 0.1, doubNextPortStartPointY + 0.9 + 0.4 * (intPortOnChassis % 2), doubNextPortStartPointX + 0.5, doubNextPortStartPointY + 1.3 + 0.4 * (intPortOnChassis % 2));
                            arrShapesLanFakeCircles[intCurrentOverallLinkNumber].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrShapesLanFakeConnections[intCurrentOverallLinkNumber] = page1.DrawRectangle(doubNextPortStartPointX + 0.3, doubNextPortStartPointY + 3.6, doubNextPortStartPointX + 0.3, doubNextPortStartPointY + 3.6);
                            doubNextPortStartPointX += 0.2;

                            // Добавление порта в MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strPortName}', {strLinkTypeId}, {strLinkSubTypeId})";
                            command.ExecuteNonQuery();
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Data1 = Convert.ToString(LastUsedId(command, "Port"));
                            //Console.WriteLine($"porta: {arrShapesLanPorts[intCurrentOverallLinkNumber].Data1}");
                        };
                    };

                    // Фиксация координаты Х - для использования при рисовании фигуры следующего устройства.
                    doubLanLastX = doubDeviceEndPointX;
                };

                // Группируем LAN устройства и порты
                foreach (Visio.Shape objSingleLanDevice in listShapesLanDevices)
                {
                    vsoWindow.DeselectAll();
                    vsoWindow.Select(objSingleLanDevice, 2);
                    for (int intCurrentLanLink = 1; intCurrentLanLink <= intTotalOverallLinkNumber; intCurrentLanLink++)
                    {
                        if (objSingleLanDevice.Data3 == arrShapesLanPorts[intCurrentLanLink].Data3)
                        {
                            vsoWindow.Select(arrShapesLanPorts[intCurrentLanLink], 2);
                            vsoWindow.Select(arrShapesLanFakeLines[intCurrentLanLink], 2);
                            vsoWindow.Select(arrShapesLanFakeCircles[intCurrentLanLink], 2);
                            vsoWindow.Select(arrShapesLanFakeConnections[intCurrentLanLink], 2);
                        };

                    };
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };

                ////////////////////////////////////////////////////////////////////////////////// Рисуем Шасси WAN ////////////////////////////////////////////////////////////////////////////

                // Сброс счётчика линков перед рисованием фигур WAN Visio
                intCurrentOverallLinkNumber = 0;

                // Счётчик слотов в стойке LAN-WAN увеличиваем на 10. Заполнение стойки LAN-устройствами было с самого низа.
                intCurrentRackSlot += 10;

                // Проход по словарю портов WAN-устройств
                foreach (Dictionary<string, object> dictWanDevices in listWanDevices)
                {
                    intCurrentWanChassis++;
                    doubDeviceStartPointX = Convert.ToDouble(dictWanDevices["StartX"]);
                    doubDeviceStartPointY = Convert.ToDouble(dictWanDevices["StartY"]);
                    doubDeviceEndPointX = Convert.ToDouble(dictWanDevices["EndX"]);
                    doubDeviceEndPointY = Convert.ToDouble(dictWanDevices["EndY"]);
                    strShapeName = $"{Convert.ToString(dictWanDevices["Device_Name"])}";
                    strRouterNameModified = Convert.ToString(dictWanDevices["Device_Name"]).Replace(" ", "_").Replace("-", "_");
                    strNameForRackTables = $"{strFullSiteIndex}-WAN-{strRouterNameModified}";
                    listShapesWanDevices.Add(page1.DrawRectangle(doubDeviceStartPointX, doubDeviceStartPointY, doubDeviceEndPointX, doubDeviceEndPointY));
                    listShapesWanDevices[listShapesWanDevices.Count - 1].Text = strShapeName;
                    listShapesWanDevices[listShapesWanDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(169,169,169)";
                    listShapesWanDevices[listShapesWanDevices.Count - 1].Data3 = Convert.ToString(intCurrentWanChassis);

                    // Добавляем шасси LAN в БД. Стойку LAN-WAN заплняем снизу вверх.
                    intCurrentRackSlot++;
                    intDeviceId = FillRackSlot(command, strObjectIndex, intLanWanRackId, intCurrentRackSlot, 50003, strNameForRackTables);

                    // Устанавливаем координаты первого порта.
                    doubNextPortStartPointX = doubDeviceStartPointX - 0.1;
                    doubNextPortStartPointY = doubDeviceStartPointY - 0.2;

                    // Проход по словарю портов WAN-устройств
                    foreach (Dictionary<string, object> dictWanPorts in listWanPorts)
                    {
                        if (Convert.ToInt32(dictWanDevices["Device_Index"]) == Convert.ToInt32(dictWanPorts["Device_Index"]))
                        {
                            switch (dictWanPorts["Link_Type"])
                            {
                                case "100G":
                                    strLinkTypeId = "15";
                                    strLinkSubTypeId = "1670";              // Переписать логику. Сделать отдельные switch case для субтайпов
                                    break;
                                case "40G":
                                    strLinkTypeId = "10";                   //Узнать!
                                    strLinkSubTypeId = "1664";
                                    break;
                                case "10G":
                                    strLinkTypeId = "9";
                                    strLinkSubTypeId = "36";
                                    break;
                                case "1G":
                                    strLinkTypeId = "4";
                                    strLinkSubTypeId = "1204";                //Узнать!
                                    break;
                            }

                            // Рисование порта с добавлением его в массив фигур LAN. Индекс порта в массиве берётся из списка словарей LAN-портов.
                            intCurrentOverallLinkNumber = Convert.ToInt32(dictWanPorts["Overall_Link_Number"]);
                            strPortName = Convert.ToString(dictWanPorts["Port_Name"]);
                            intPortOnChassis++;
                            arrShapesWanPorts[intCurrentOverallLinkNumber] = page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.6, doubNextPortStartPointY - 0.2);
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Data1 = Convert.ToString(dictWanPorts["Port_Index"]);
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Data2 = Convert.ToString(dictWanPorts["Link_Type"]);
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Data3 = Convert.ToString(intCurrentWanChassis);
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Text = strPortName;
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Rotate90();
                            arrShapesWanPorts[intCurrentOverallLinkNumber].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            arrShapesWanFakeLines[intCurrentOverallLinkNumber] = page1.DrawLine(doubNextPortStartPointX + 0.3, doubNextPortStartPointY - 3.6, doubNextPortStartPointX + 0.3, doubNextPortStartPointY - 0.4);
                            arrShapesWanFakeCircles[intCurrentOverallLinkNumber] = page1.DrawOval(doubNextPortStartPointX + 0.1, doubNextPortStartPointY - 1.1 - 0.4 * (intPortOnChassis % 2), doubNextPortStartPointX + 0.5, doubNextPortStartPointY - 0.7 - 0.4 * (intPortOnChassis % 2));
                            arrShapesWanFakeCircles[intCurrentOverallLinkNumber].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrShapesWanFakeConnections[intCurrentOverallLinkNumber] = page1.DrawRectangle(doubNextPortStartPointX + 0.3, doubNextPortStartPointY - 3.6, doubNextPortStartPointX + 0.3, doubNextPortStartPointY - 3.6);
                            doubNextPortStartPointX += 0.2;

                            // Добавление порта в MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strPortName}', {strLinkTypeId}, {strLinkSubTypeId})";
                            command.ExecuteNonQuery();
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Data1 = Convert.ToString(LastUsedId(command, "Port"));
                        };
                    };
                    doubWanLastX = doubDeviceEndPointX;
                };

                //Фиксация точки Y, с которой рисовать следующие устройства.
                double intdoubTopLineY = 60;
                intGlobalCableCounter = 0;

                //Группируем WAN устройства и порты
                foreach (Visio.Shape objSingleWanDevice in listShapesWanDevices)
                {
                    vsoWindow.DeselectAll();
                    vsoWindow.Select(objSingleWanDevice, 2);
                    for (int intCurrentWanLink = 1; intCurrentWanLink <= intTotalOverallLinkNumber; intCurrentWanLink++)
                    {
                        if (objSingleWanDevice.Data3 == arrShapesLanPorts[intCurrentWanLink].Data3)
                        {
                            vsoWindow.Select(arrShapesWanPorts[intCurrentWanLink], 2);
                            vsoWindow.Select(arrShapesWanFakeLines[intCurrentWanLink], 2);
                            vsoWindow.Select(arrShapesWanFakeCircles[intCurrentWanLink], 2);
                            vsoWindow.Select(arrShapesWanFakeConnections[intCurrentWanLink], 2);
                        };

                    };
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };

                // Группируем WAN устройства и порты
                List<Visio.Shape> listDeviceMgmtPorts = new List<Visio.Shape>();                                                                //Add MGMT Ports List (for all TSPU devices)
                List<Visio.Shape> listDeviceMgmtLines = new List<Visio.Shape>();
                List<Visio.Shape> listDeviceMgmtFakeRects = new List<Visio.Shape>();
                List<Visio.Shape> listDeviceMgmtCircles = new List<Visio.Shape>();


                ////////////////////////////////////////////////////////////////////////////////// Рисуем Байпасы ////////////////////////////////////////////////////////////////////////////

                // Подсчёт количества линков на один балансер при прямом включении десяток
                //int LastPortOnBalancer = CalculateDevicesQuantity(intLinkCounter10, Convert.ToInt32(strBalancerNumberFromInput));
                //int LastPortOnBalancer = 33;

                // Переменная для подсчёта количества портов MGMT на IBS1UP.
                int intDifference;

                // Переменные первых координат байпасов.
                double doubUpperStartPoint = doubStartPointNextShapeY - 1.5;
                double doubFirstBypassPortY;
                double doubLastBypassPortY;

                // Создаём список фигур шасси IS100.
                List<Visio.Shape> listShapesBypass100Devices = new List<Visio.Shape>();


                //~~~~~~~~~~~~~~~~~~~~~~~~~~~   Signaltek   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                // Создаём список заголовков фигур шасси Signaltek.
                //List<Visio.Shape> listShapesSignaltekHeaders = new List<Visio.Shape>();
                Visio.Shape[] listShapesSignaltekHeaders = new Visio.Shape[50];

                // Создаём массив модулей шасси Сигналтек.
                Visio.Shape[,] arrShapesSignaltekModules = new Visio.Shape[50, 4];

                // Создаём массив модулей портов Сигналтек.
                Visio.Shape[,,] arrShapesSignaltekNetPorts = new Visio.Shape[50, 4, 2];
                Visio.Shape[,,] arrShapesSignaltekMonPorts = new Visio.Shape[50, 4, 2];
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                // Создаём массив фигур портов IS100.
                Visio.Shape[,] arrShapesBypass100MonPorts = new Visio.Shape[10, 200];
                Visio.Shape[,] arrShapesBypass100NetPorts = new Visio.Shape[200, 10];

                // Создаём массив фигур портов IS40/IBS1UP (1G/10G): Net, Mon, подпись ODF.
                Visio.Shape[,] arrShapesBypass10_NetPorts = new Visio.Shape[200, 10];
                Visio.Shape[,,] arrShapesBypass10_MonPorts = new Visio.Shape[200, 200, 10];
                Visio.Shape[,] arrShapesBypass10_Odf = new Visio.Shape[200, 10];

                // Создаём массив фигур портов IS40 (40G).
                Visio.Shape[,,] arrShapesBypass40_MonPorts = new Visio.Shape[200, 200, 10];
                Visio.Shape[,] arrShapesBypass40_NetPorts = new Visio.Shape[200, 10];
                Visio.Shape[,] arrShapesBypass40_Odf = new Visio.Shape[200, 10];

                // Данные для группировок портов. Возможно, удалить.
                int int100mDevice1 = 0;
                int int100mPort1 = 0;
                int int100mDevice2 = 0;
                int int100mPort2 = 0;
                int int100mDevice3 = 0;
                int int100mPort3 = 0;
                int int100mDevice4 = 0;
                int int100mPort4 = 0;

                // Создаём массив фигур портов и гидр фильтров.         (Уточнить!)
                Visio.Shape[,] arrShapesFilterPorts = new Visio.Shape[200, 20];
                Visio.Shape[,] arrFilterHydraConnectors = new Visio.Shape[200, 10];
                Visio.Shape[,] arrBypassIs40HydraConnectors = new Visio.Shape[200, 200];

                // Создаём массив фигур Net-портов байпасов.
                Visio.Shape[,] arrShapesBypassNetFakeLine = new Visio.Shape[200, 10];
                Visio.Shape[,] arrShapesBypassNetFakeConnection = new Visio.Shape[200, 10];
                Visio.Shape[,] arrShapesBypassNetFakeCircles = new Visio.Shape[200, 10];

                // Создаём массив фигур Mon-портов байпасов 100G.
                Visio.Shape[,] arrShapesBypass100MonFakeLine = new Visio.Shape[200, 50];
                Visio.Shape[,] arrShapesBypass100MonFakeConnection = new Visio.Shape[200, 50];
                Visio.Shape[,] arrShapesBypassMonFakeCircles = new Visio.Shape[200, 50];                // Разобраться с переиспользованием фигур для 10G и 100G

                // Создаём массив фигур Mon-портов байпасов 10G.
                Visio.Shape[,,] arrShapesBypass10MonFakeLine = new Visio.Shape[200, 50, 10];
                Visio.Shape[,,] arrShapesBypass10MonFakeConnection = new Visio.Shape[200, 50, 10];

                // Переменные с текущими именами девайса и порта
                string strCurrentDeviceHostname;
                string strCurrentPortName;

                // Переменная правой координаты Х байпасов. Для расчёта координат устройств правее ряда байпасов.
                double doubBypassEndX;

                // Переменная для рассчёта координаты Y палки ODF.
                doubFirstBypassPortY = doubStartPointNextShapeY - 1.3;

                // Расчёт левой координы Х байпасов. Из количества линков - из длины LAN-WAN.
                doubStartPointNextShapeX = intTotalOverallLinkNumber * 2 + 6;           //Возможно, изменить
                doubStartPointNextShapeY -= 2.5;
                double doubFirstBypassPortX = doubStartPointNextShapeX - 5.2;

                // Переменные счётчиков портов байпаса.
                int intCurrentBypassPort = 0;
                int intCurrentBalancerChassis = 0;

                // Переменная массива количества портов на каждом из балансеров.
                int[] arrCurrentUplinkPortInBalancer = new int[20];

                // Переменная принадлежности порта старому лагу либо новому.
                bool boolNotEnoughPortsForLag = false;

                // Переменная наличия не более двух десяток в последней гидре трансграна.
                //bool noMoreTwoLinksInHydra = false;


                //int intPortsTaken = 32;

                // Обнуляем поинтеры для каждого балансировщика
                for (int intCurrentBalancer = 1; intCurrentBalancer <= intTotalBalancers; intCurrentBalancer++)
                {
                    arrCurrentUplinkPortInBalancer[intCurrentBalancer] = 16;
                }

                // Обнуление счётчика линков и портов балансера.
                intCurrentOverallLinkNumber = 0;
                intCurrent100mBalancerPort = 0;

                // Установка начальной позиции деваймов в рэке (MySQL).
                intCurrentRackSlot = 10;

                // Переменные ID портов - текущего и соседнего с ним (MySQL).
                string strNeighborPortId;
                string strLocalPortId;

                //// Рисуем IS100.
                if (boolEshelon) Console.WriteLine("Рисуем Эшелон");                    // Уведомление о том, что схема эшелонная.
                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs100Bypasses; intCurrentBypassDevice++)
                {
                    listShapesBypass100Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 1.3));
                    strCurrentDeviceHostname = "IS100 (" + intCurrentBypassDevice + ")";
                    strNameForRackTables = $"{strFullSiteIndex}-IS100G-{intCurrentBypassDevice}";
                    listShapesBypass100Devices[listShapesBypass100Devices.Count - 1].Text = strCurrentDeviceHostname;
                    listShapesBypass100Devices[listShapesBypass100Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(255,228,225)";
                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGNT ETH";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MGNT ETH";
                    //listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(intCurrentBypassDevice);

                    //Mgmt Lines
                    listDeviceMgmtLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5));
                    listDeviceMgmtCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.7, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.1));
                    listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                    listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                    listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5));
                    listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "MGNT ETH";

                    // Установка значений координат первых портов байпаса, чьё шасси только что нарисовали.
                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    // Установка координаты Y для следующего шасси байпаса.
                    doubStartPointNextShapeY -= 6;

                    // Добавление шасси в БД.
                    intCurrentRackSlot++;
                    intDeviceId = FillRackSlot(command, strObjectIndex, intBypassRackId, intCurrentRackSlot, 50001, strNameForRackTables);

                    // MySQL Mgmt.
                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'MGNT ETH', 1, 24);";
                    command.ExecuteNonQuery();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                    // Рисуем порты IS100.
                    // Рисуем Net-порты.
                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 2; intCurrentPortCounterInBypass++)
                    {
                        // Сдвигаем итераторы общего количества линков и портов текущего байпаса.
                        intCurrentOverallLinkNumber++;
                        intCurrentBypassPort++;


                        // Отрисовка портов Net0
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/0";
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].Data3 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].Text = strCurrentPortName;
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrent100mBalancerPort++;
                        if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                        {
                            // Odf Text
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 3.6, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 2.6, doubNextPortStartPointY - 0.3);
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].Text = arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Side_A_Port"];
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].LineStyle = "None";
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].Style = "None";
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            // Fakes
                            arrShapesBypassNetFakeLine[intCurrentOverallLinkNumber, 0] = page1.DrawLine(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4);
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0] = page1.DrawOval(doubNextPortStartPointX - 1.1, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.7, doubNextPortStartPointY - 0.6);
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0].Text = intCurrentOverallLinkNumber + " L";
                            arrShapesLanFakeCircles[intCurrentOverallLinkNumber].Text = intCurrentOverallLinkNumber + " L";
                            arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4);
                            // Connect
                            arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 0].AutoConnect(arrShapesLanFakeConnections[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            // Дозаполнение ячейки массива словарей LAN-Bypass
                            if (arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Comment"] == "Прямое включение") SideADevice = arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Side_A_Name"];
                            else SideADevice = "ODF";
                            arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Name", strCurrentDeviceHostname);
                            arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Port", strCurrentPortName);
                            //arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Name", $"ODF --- {strBypassModel} ({intCurrentBypassDevice})");
                            arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Name", $"{SideADevice} --- {strBypassModel} ({intCurrentBypassDevice})\n100GBASE-LR4");
                            arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Number", intCurrentOverallLinkNumber + " L");
                            list_CableJournal_LAN_Bypass.Add(new Dictionary<string, string>());
                            // Перенос словарей из нумерованного списка в ненумерованный - чтобы отработал foreach
                            foreach (string key in arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Keys)
                            {
                                list_CableJournal_LAN_Bypass[list_CableJournal_LAN_Bypass.Count - 1].Add(key, arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0][key]);
                            };
                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 15, 1670);";
                            command.ExecuteNonQuery();
                            strNeighborPortId = arrShapesLanPorts[intCurrentOverallLinkNumber].Data1;
                            strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                            command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                            command.ExecuteNonQuery();
                        };

                        // Отрисовка портов Net1
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/1";
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].Data3 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].Text = strCurrentPortName;
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrent100mBalancerPort++;
                        if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                        {
                            // Odf Text
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 3.6, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 2.6, doubNextPortStartPointY - 0.1);
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].Text = arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Side_A_Port"];
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].LineStyle = "None";
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].Style = "None";
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            // Fakes
                            arrShapesBypassNetFakeLine[intCurrentOverallLinkNumber, 1] = page1.DrawLine(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.2);
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1] = page1.DrawOval(doubNextPortStartPointX - 1.6, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 1.2, doubNextPortStartPointY);
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1].Text = intCurrentOverallLinkNumber + " W";
                            arrShapesWanFakeCircles[intCurrentOverallLinkNumber].Text = intCurrentOverallLinkNumber + " W";
                            arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2);
                            //Connect
                            arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 1].AutoConnect(arrShapesWanFakeConnections[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            //Дозаполнение ячейки массива словарей WAN-Bypass
                            if (arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Comment"] == "Прямое включение") SideADevice = arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Side_A_Name"];
                            else SideADevice = "ODF";
                            arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Name", strCurrentDeviceHostname);
                            arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Port", strCurrentPortName);
                            //arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Name", $"ODF --- {strBypassModel} ({intCurrentBypassDevice})");
                            arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Name", $"{SideADevice} --- {strBypassModel} ({intCurrentBypassDevice})\n100GBASE-LR4");
                            arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Number", intCurrentOverallLinkNumber + " W");
                            list_CableJournal_WAN_Bypass.Add(new Dictionary<string, string>());
                            foreach (string key in arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Keys)
                            {
                                list_CableJournal_WAN_Bypass[list_CableJournal_WAN_Bypass.Count - 1].Add(key, arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1][key]);
                            };
                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 15, 1670);";
                            command.ExecuteNonQuery();
                            strNeighborPortId = arrShapesWanPorts[intCurrentOverallLinkNumber].Data1;
                            strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                            command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                            command.ExecuteNonQuery();
                        };

                        // Рисуем Mon-порты.
                        // Удалить!
                        if (intCurrentPortCounterInBypass == 1) list_Specification.Add("IS100M100G4BP-QL4S4");
                        list_Specification.Add("NR-QSFP-100G-AOC 100G QSFP28 <=> QSFP28");
                        list_Specification.Add("NR-QSFP-100G-AOC 100G QSFP28 <=> QSFP28");
                        //Раундробин (формула).
                        intCurrentBalancerChassis++;
                        //WAN-линки - в нечётные порты балансировщика (1, 3, 5, 7,...)
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/1";
                        intCurrentPortInChassis = arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 1;
                        // Запись номеров шасси и портов
                        if (intCurrentPortCounterInBypass == 1)
                        {
                            int100mDevice1 = intCurrentBalancerChassis;
                            int100mPort1 = intCurrentPortInChassis;
                        }
                        else
                        {
                            int100mDevice2 = intCurrentBalancerChassis;
                            int100mPort2 = intCurrentPortInChassis;
                        };
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data3 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Text = strCurrentPortName;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        // MySQL
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 15, 1672);";
                        command.ExecuteNonQuery();
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(LastUsedId(command, "Port"));
                        //КЖ линков WAN
                        intGlobalCableCounter++;
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0] = new Dictionary<string, string>();
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + " ББ");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Name", $"{strCurrentDeviceHostname} --- ELB-0133 ({intCurrentBalancerChassis})");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Type", "NR-QSFP-100G-AOC 100G QSFP28 <=> QSFP28");
                        //Draw Line And Fake Rectangle
                        arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.2);
                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawOval(doubNextPortStartPointX + 2.9, doubNextPortStartPointY, doubNextPortStartPointX + 3.3, doubNextPortStartPointY - 0.4);
                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].Text = Convert.ToString(intGlobalCableCounter) + " ББ";
                        arrShapesBypass100MonFakeConnection[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.2);
                        //LAN-линки - в чётные порты балансировщика (2, 4, 6, 8,...)
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                        if (!boolEshelon) intCurrentPortInChassis = arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 2;
                        else intCurrentPortInChassis = arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 9;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data3 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Text = strCurrentPortName;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        // MySQL
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 15, 1672);";
                        command.ExecuteNonQuery();
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(LastUsedId(command, "Port"));
                        //КЖ линков LAN
                        intGlobalCableCounter++;
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0] = new Dictionary<string, string>();
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + " ББ");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Name", $"{strCurrentDeviceHostname} --- ELB-0133 ({intCurrentBalancerChassis})");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Type", "NR-QSFP-100G-AOC 100G QSFP28 <=> QSFP28");
                        //Draw Line And Fake Rectangle
                        arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.4);
                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawOval(doubNextPortStartPointX + 3.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 3.9, doubNextPortStartPointY - 0.6);
                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].Text = Convert.ToString(intGlobalCableCounter) + " ББ";
                        arrShapesBypass100MonFakeConnection[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.4);
                        //После отрисовки пары WAN-LAN сдвиг указателя на 2 порта для нижнего эшелона и на один порт для верхнего.
                        if (!boolEshelon) arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] += 2;
                        else arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] += 1;
                        //После полного прогона балансировщиков указатель перемещается в начало - к первому балансировщику.
                        if (intCurrentBalancerChassis == intTotalBalancers) intCurrentBalancerChassis = 0;
                        //Сдвиг вниз следующей фигуры байпаса
                        doubNextPortStartPointY -= 0.7;
                    };
                };

                // Фиксация количества портов балансировщика при включении 100G.
                intTotal100mBalancerPorts = intCurrent100mBalancerPort;

                // Для 100G отсутствует необходимость работы с лагами. Зануляем значения.
                intCurrentLagNumber = 1;
                arrLagIteration[intCurrentLagNumber] = 0;


                //Создание фигур для отрисовки IBS1UP/IS40 (для 10G)
                List<Visio.Shape> listShapesBypass10Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass10NetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass10MonPorts = new List<Visio.Shape>();
                List<Visio.Shape> listHydraLines = new List<Visio.Shape>();
                List<Visio.Shape> listEshelonWanHydraLines = new List<Visio.Shape>();
                List<Visio.Shape> listEshelonLanHydraLines = new List<Visio.Shape>();
                List<Visio.Shape> listBypassHydraConnectors = new List<Visio.Shape>();
                List<Visio.Shape> listShapesIbs1UpDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesIbs1upNetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesIbs1upMonPorts = new List<Visio.Shape>();

                // Вспомогательные переменные для 10G байпасов. 
                int iCurrentOverallMonPort = 0;
                int intFilterPortNoBalancer = intPortsNumberOnSingleFilter / 2;
                int intBypassIs40CurrentHydra = 0;
                int intBypassIs40HydrasTotal = 0;
                bool boolLastBypassChassis = false;
                int intMgmtPortOnChassis;
                int intHydraEnd;
                string strHydraName;
                int intEshelonBalancerPortShift = 0;
                intCurrentBalancerChassis = 0;
                intCurrentPortInChassis = 16;
                intCurrentOverallLinkNumber = 0;
                int intFirstBaypassChassis;
                int intLastBypassChassis;
                bool boolFinishedFillingSingleBalancer = true;

                // Переменные для фиксации координиат лучей гидр эшелонной схемы.
                double doubLongWanHydraStartCoordinateY = 0;
                double doubLongLanHydraStartCoordinateY = 0;

                // Сдвиг ряда фигур IS40 вниз
                doubStartPointNextShapeY -= 1;

                // Обнуление счётчика портов байпаса
                intCurrentBypassPort = 0;

                // Подсчёт количества занятых байпас-сегментов в IBS1UP. Чтобы завявить меньше MGMT-портов на шасси.
                intDifference = intTotalOverallLinkNumber % 4;
                Console.WriteLine($"Остаток от деления: {intDifference}");
                //Console.WriteLine($"Отрисовка байпасов начата.");

                // Проход по всем десяточным байпасам (IBS1UP).
                int intCurrentFilterNoBalancer = 0;

                // Сброс переменной ID девайса (MySQL).
                intDeviceId = 0;

                // Рисование IBS1UP или IS40. Выбор в зависимости от значения strBypassModel.
                // Фиксация значений переменных в зависимости от того, какое сочетание байпасов: IS40, IBS1UP или MIX.
                if (strBypassModel == "IBS1UP" || strBypassModel == "MIX")
                {
                    if (strBypassModel == "MIX")
                    {
                        //Console.WriteLine($"Миксовая схема.");
                        intLastBypassChassis = intBypasses10Old;
                    }
                    else
                    {
                        //Console.WriteLine($"Немиксовая схема.");
                        intLastBypassChassis = intTotalIs10Bypasses;
                    };
                    //Console.WriteLine($"Заполняем IBS1UP.");
                    //Console.WriteLine($"intLastBypassChassis = {intLastBypassChassis}");

                    // Рисуем шасси IBS1UP.
                    for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intLastBypassChassis; intCurrentBypassDevice++)
                    {
                        // Рисуем шасси IBS1UP.
                        listShapesIbs1UpDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 2.7));
                        strCurrentDeviceHostname = "IBS1UP (" + intCurrentBypassDevice + ")";
                        strNameForRackTables = $"{strFullSiteIndex}-IBS1UP-{intCurrentBypassDevice}";
                        listShapesIbs1UpDevices[listShapesIbs1UpDevices.Count - 1].Text = strCurrentDeviceHostname;
                        listShapesIbs1UpDevices[listShapesIbs1UpDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(135,206,250)";
                        listShapesIbs1UpDevices[listShapesIbs1UpDevices.Count - 1].Data3 = strCurrentDeviceHostname;

                        // Добавляем шасси в БД (MySQL).
                        intCurrentRackSlot++;
                        intDeviceId = FillRackSlot(command, strObjectIndex, intBypassRackId, intCurrentRackSlot, 50001, strNameForRackTables);

                        // Сдуигаем координаты порто
                        double doubMgmtPortX = doubStartPointNextShapeX + 0.3;
                        double doubMgmtPortY = doubStartPointNextShapeY + 0.15;

                        //В миксовой схеме: если текущее шасси соответствует последнему шасси IBS1UP, ставим соответствующий флаг.
                        if (intCurrentBypassDevice == intLinkCounter10Old || intCurrentBypassDevice == intLastBypassChassis) boolLastBypassChassis = true;

                        //Если последнее шасси IBS1UP, то количество портов MGMT считаем из разницы между 4 и оставшимися портами
                        //if (boolLastBypassChassis) intMgmtPortOnChassis = 4 - intDifference;
                        if (boolLastBypassChassis) intMgmtPortOnChassis = intDifference;
                        else intMgmtPortOnChassis = 4;


                        //intMgmtPortOnChassis = 3;                                             // Количество MGMT-портов IBS1UP
                        // Рисуем порты Mgmt на шасси.
                        for (int intIbs1upCurrentPortMgmt = 1; intIbs1upCurrentPortMgmt <= intMgmtPortOnChassis; intIbs1upCurrentPortMgmt++)
                        {
                            listDeviceMgmtPorts.Add(page1.DrawRectangle(doubMgmtPortX, doubMgmtPortY, doubMgmtPortX + 0.5, doubMgmtPortY + 0.2));
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGMT-" + intIbs1upCurrentPortMgmt;
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MANAGEMENT-" + intIbs1upCurrentPortMgmt;
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = strCurrentDeviceHostname;
                            listDeviceMgmtLines.Add(page1.DrawLine(doubMgmtPortX + 0.25, doubMgmtPortY + 0.35, doubMgmtPortX + 0.25, doubMgmtPortY + 1.5 + 0.2 * (4 - intIbs1upCurrentPortMgmt)));
                            listDeviceMgmtCircles.Add(page1.DrawOval(doubMgmtPortX + 0.05, doubMgmtPortY + 0.5 + 0.4 * (intIbs1upCurrentPortMgmt % 2), doubMgmtPortX + 0.45, doubMgmtPortY + 0.9 + 0.4 * (intIbs1upCurrentPortMgmt % 2)));
                            listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                            listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubMgmtPortX + 0.25, doubMgmtPortY + 1.5 + 0.2 * (4 - intIbs1upCurrentPortMgmt), doubMgmtPortX + 0.25, doubMgmtPortY + 1.5 + 0.2 * (4 - intIbs1upCurrentPortMgmt)));
                            listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                            listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "MANAGEMENT-" + intIbs1upCurrentPortMgmt;
                            listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data3 = strCurrentDeviceHostname;

                            // MySQL Mgmt.
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'MANAGEMENT-{intIbs1upCurrentPortMgmt}', 1, 24);";
                            command.ExecuteNonQuery();
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                            //Сдвиг координаты Х вправо
                            doubMgmtPortX += 0.2;
                        };



                        // Вычисление координат следующего шасси IBS1UP.
                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;
                        doubStartPointNextShapeY -= 8;

                        // Рисуем порты IBS1UP (10G).
                        for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 4; intCurrentPortCounterInBypass++)
                        {

                            //Net-порты
                            intCurrentOverallLinkNumber++;
                            intCurrentBypassPort++;


                            //NetX/0
                            strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/0";
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Text = strCurrentPortName;
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                            {
                                list_Specification.Add("duplex LC/UPC-LC/UPC, SM");
                                //Odf Text
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 3.6, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 2.6, doubNextPortStartPointY - 0.3);
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].Text = arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Side_A_Port"];
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].LineStyle = "None";
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].Style = "None";
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                //Fakes
                                arrShapesBypassNetFakeLine[intCurrentOverallLinkNumber, 0] = page1.DrawLine(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4);
                                arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0] = page1.DrawOval(doubNextPortStartPointX - 1.1, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.7, doubNextPortStartPointY - 0.6);
                                arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0].Text = intCurrentOverallLinkNumber + " L";
                                arrShapesLanFakeCircles[intCurrentOverallLinkNumber].Text = intCurrentOverallLinkNumber + " L";
                                arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4);
                                //Connect
                                arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 0].AutoConnect(arrShapesLanFakeConnections[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                //Дозаполнение ячейки массива словарей LAN-Bypass
                                if (arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Comment"] == "Прямое включение") SideADevice = arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Side_A_Name"];
                                else SideADevice = "ODF";
                                arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Name", strCurrentDeviceHostname);
                                arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Port", strCurrentPortName);
                                arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Number", intCurrentOverallLinkNumber + " L");
                                arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Name", $"{SideADevice} --- {strBypassModel} ({intCurrentBypassDevice})\n10GBASE-LR");
                                //arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Name", $"ODF --- {strBypassModel} ({intCurrentBypassDevice})");
                                list_CableJournal_LAN_Bypass.Add(new Dictionary<string, string>());
                                foreach (string key in arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Keys)
                                {
                                    list_CableJournal_LAN_Bypass[list_CableJournal_LAN_Bypass.Count - 1].Add(key, arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0][key]);
                                };
                                // MySQL
                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', {strLinkTypeId}, {strLinkSubTypeId});";
                                command.ExecuteNonQuery();
                                strNeighborPortId = arrShapesLanPorts[intCurrentOverallLinkNumber].Data1;
                                strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                                command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                                command.ExecuteNonQuery();

                            };

                            //NetX/1
                            strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/1";
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data2 = Convert.ToString(intCurrentBypassPort);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Text = strCurrentPortName;
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                            {
                                list_Specification.Add("duplex LC/UPC-LC/UPC, SM");
                                // Odf Text
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 3.6, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 2.6, doubNextPortStartPointY - 0.1);
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].Text = arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Side_A_Port"];
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].LineStyle = "None";
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].Style = "None";
                                arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                // Рисование линий и кружков.
                                arrShapesBypassNetFakeLine[intCurrentOverallLinkNumber, 1] = page1.DrawLine(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.2);
                                arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1] = page1.DrawOval(doubNextPortStartPointX - 1.6, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 1.2, doubNextPortStartPointY);
                                arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1].Text = intCurrentOverallLinkNumber + " W";
                                arrShapesWanFakeCircles[intCurrentOverallLinkNumber].Text = intCurrentOverallLinkNumber + " W";
                                arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2);
                                // Connect
                                arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 1].AutoConnect(arrShapesWanFakeConnections[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                // Дозаполнение ячейки массива словарей WAN-Bypass
                                if (arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Comment"] == "Прямое включение") SideADevice = arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Side_A_Name"];
                                else SideADevice = "ODF";
                                arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Name", strCurrentDeviceHostname);
                                arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Port", strCurrentPortName);
                                arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Number", intCurrentOverallLinkNumber + " W");
                                //arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Name", $"ODF --- {strBypassModel} ({intCurrentBypassDevice})");
                                arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Name", $"{SideADevice} --- {strBypassModel} ({intCurrentBypassDevice})\n10GBASE-LR");
                                list_CableJournal_WAN_Bypass.Add(new Dictionary<string, string>());
                                foreach (string key in arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Keys)
                                {
                                    list_CableJournal_WAN_Bypass[list_CableJournal_WAN_Bypass.Count - 1].Add(key, arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1][key]);
                                };
                                // MySQL
                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', {strLinkTypeId}, {strLinkSubTypeId});";
                                command.ExecuteNonQuery();
                                strNeighborPortId = arrShapesWanPorts[intCurrentOverallLinkNumber].Data1;
                                strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                                command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                                command.ExecuteNonQuery();
                            };
                            // Отслеживание лагов
                            arrLagIteration[intCurrentLagNumber]++;
                            if (arrLagIteration[intCurrentLagNumber] == 1)
                            {
                                if (32 - intCurrentPortInChassis < arrLanLagCounter[intCurrentLagNumber] / 2)
                                {
                                    boolNotEnoughPortsForLag = true;
                                };
                            };
                            // Стык байпаса напрямую к фильтру.
                            if (!boolNoBalancer)
                            {
                                if (listHydraLines.Count == 0) intGlobalCableCounter++;                      // Проверить!
                                if ((intCurrentBypassPort - 1) % 2 == 0) intCurrentPortInChassis++;
                                if (intCurrentPortInChassis == 33 || intCurrentBalancerChassis == 0 || boolNotEnoughPortsForLag)
                                {
                                    if (intCurrentBalancerChassis > 0)
                                    {
                                        arrUplinkPortsOnBalancer[intCurrentBalancerChassis] = intCurrentPortInChassis - 1;
                                        //Console.WriteLine($"На балансировщике {intCurrentBalancerChassis} последний заполненный порт: {intCurrentPortInChassis - 1}.");
                                    };
                                    boolNotEnoughPortsForLag = false;
                                    intCurrentBalancerChassis++;
                                    intCurrentPortInChassis = 17;
                                };
                                if (arrLagIteration[intCurrentLagNumber] == arrLanLagCounter[intCurrentLagNumber])
                                {
                                    intCurrentLagNumber++;
                                    arrLagIteration[intCurrentLagNumber] = 0;
                                };
                            };
                            if (intFilterPortNoBalancer == intPortsNumberOnSingleFilter / 2)
                            {
                                intCurrentFilterNoBalancer++;
                                intFilterPortNoBalancer = 0;
                            };

                            // Mon X/1
                            iCurrentOverallMonPort++;
                            strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/1";
                            // Стык порта байпаса к порту балансировщика.
                            if (!boolNoBalancer)
                            {
                                if (intCurrentPortCounterInBypass == 1 || intCurrentPortCounterInBypass == 3)
                                {
                                    strCableInHydra = " (AOC c1)";
                                    intHydraEnd = 1;
                                }
                                else
                                {
                                    strCableInHydra = " (AOC c3)";
                                    intHydraEnd = 3;
                                };
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 1] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 1].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 1].Text = strCurrentPortName;
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                // КЖ.
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd] = new Dictionary<string, string>();
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Row", Convert.ToString(intCurrentJournalItem));
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Name", $"{strCurrentDeviceHostname} --- ELB-0133 ({intCurrentBalancerChassis})");
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + strCableInHydra);
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Device_A_Name", strCurrentDeviceHostname);
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Type", "NR-QSFP-4X10G-AOC 40G QSFP");
                                // MySQL
                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                //command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', {strLinkTypeId}, {strLinkSubTypeId});";
                                command.ExecuteNonQuery();
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));
                            }
                            // Стык порта байпаса к порту фильтра.
                            else if (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                            {
                                intFilterPortNoBalancer++;
                                strCableInHydra = "";
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Text = strCurrentPortName;
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                // Рисование линий и кружков.
                                arrShapesBypass100MonFakeLine[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1] = page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 4, doubNextPortStartPointY - 0.2);
                                arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1] = page1.DrawOval(doubNextPortStartPointX + 3.2, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 3.6, doubNextPortStartPointY);
                                arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1].Text = intFilterPortNoBalancer * 2 - 1 + " БФ";
                                arrShapesBypass100MonFakeConnection[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1] = page1.DrawRectangle(doubNextPortStartPointX + 4, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 4, doubNextPortStartPointY - 0.2);
                                // КЖ.
                                intGlobalCableCounter++;
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0] = new Dictionary<string, string>();
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + " БФ");
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Name", $"{strCurrentDeviceHostname} --- {strFilterModel} ({intCurrentFilterNoBalancer})");
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Type", "NR-SFP-10G-AOC 10G SFP+ <=> SFP+");
                                // MySQL.
                                //command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', {strLinkTypeId}, {strLinkSubTypeId});";
                                command.ExecuteNonQuery();
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));
                            };

                            // Mon X/0
                            iCurrentOverallMonPort++;
                            strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                            // Стык порта байпаса к порту балансировщика.
                            if (!boolNoBalancer)
                            {
                                if (intCurrentPortCounterInBypass == 1 || intCurrentPortCounterInBypass == 3)
                                {
                                    strCableInHydra = " (AOC c2)";
                                    intHydraEnd = 2;
                                }
                                else
                                {
                                    strCableInHydra = " (AOC c4)";
                                    intHydraEnd = 4;
                                };
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 0] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Text = strCurrentPortName;
                                arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd] = new Dictionary<string, string>();
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Row", Convert.ToString(intCurrentJournalItem));
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Name", $"{strCurrentDeviceHostname} --- ELB-0133 ({intCurrentBalancerChassis})");
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + strCableInHydra);
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Device_A_Name", strCurrentDeviceHostname);
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Type", "NR-QSFP-4X10G-AOC 40G QSFP");
                                //////////  MySQL
                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                //command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', {strLinkTypeId}, {strLinkSubTypeId});"; 
                                command.ExecuteNonQuery();
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));
                            }
                            // Стык порта байпаса к порту фильтра.
                            else if (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                            {
                                strCableInHydra = "";
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Text = strCurrentPortName;
                                arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                //Draw Line And Fake Rectangle
                                arrShapesBypass100MonFakeLine[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2] = page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 3.6, doubNextPortStartPointY - 0.4);
                                arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2] = page1.DrawOval(doubNextPortStartPointX + 2.8, doubNextPortStartPointY - 0.6, doubNextPortStartPointX + 3.2, doubNextPortStartPointY - 0.2);
                                arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2].Text = intFilterPortNoBalancer * 2 + " БФ";
                                arrShapesBypass100MonFakeConnection[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2] = page1.DrawRectangle(doubNextPortStartPointX + 3.6, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 3.6, doubNextPortStartPointY - 0.4);
                                intGlobalCableCounter++;
                                // КЖ.
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0] = new Dictionary<string, string>();
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + " БФ");
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Cable_Name", $"{strCurrentDeviceHostname} --- {strFilterModel} ({intCurrentFilterNoBalancer})");
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Cable_Type", "NR-SFP-10G-AOC 10G SFP+ <=> SFP+");
                                // MySQL
                                //command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', {strLinkTypeId}, {strLinkSubTypeId});";
                                command.ExecuteNonQuery();
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));
                            };
                            switch (intCurrentPortCounterInBypass)
                            {
                                case 1:
                                    int100mDevice1 = intCurrentFilterNoBalancer;
                                    int100mPort1 = intFilterPortNoBalancer * 2 - 1;
                                    break;
                                case 2:
                                    int100mDevice2 = intCurrentFilterNoBalancer;
                                    int100mPort2 = intFilterPortNoBalancer * 2 - 1;
                                    break;
                                case 3:
                                    int100mDevice3 = intCurrentFilterNoBalancer;
                                    int100mPort3 = intFilterPortNoBalancer * 2 - 1;
                                    break;
                                case 4:
                                    int100mDevice4 = intCurrentFilterNoBalancer;
                                    int100mPort4 = intFilterPortNoBalancer * 2 - 1;
                                    break;
                            };
                            // Рисуем коннектор гидры.
                            // Если стык с балансером и мы не вышли за пределы линков.
                            if (!boolNoBalancer && (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)) boolFinishedFillingSingleBalancer = true;

                            //boolFinishedFillingSingleBalancer
                            if (!boolNoBalancer && (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2))
                            {
                                listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.4));
                                listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.2));
                                // Если нарисованы 4 коннектора гидра.
                                if (listHydraLines.Count == 4)
                                {
                                    list_Specification.Add("NR-QSFP-4X10G-AOC 40G QSFP");  // Удалить!
                                    listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY + 0.5));
                                    vsoWindow.DeselectAll();
                                    foreach (Visio.Shape objHydraSingleLine in listHydraLines)
                                    {
                                        vsoWindow.Select(objHydraSingleLine, 2);
                                    };
                                    vsoSelection = vsoWindow.Selection;
                                    arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis] = vsoSelection.Group();
                                    arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(intCurrentBypassDevice);
                                    arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis].Data3 = Convert.ToString(intCurrentBypassDevice);
                                    intBypassIs40CurrentHydra++;
                                    listHydraLines.Clear();

                                    //Draw Line And Fake Rectangle
                                    arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawLine(doubNextPortStartPointX + 2.6, doubNextPortStartPointY, doubNextPortStartPointX + 5.6, doubNextPortStartPointY);
                                    arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawOval(doubNextPortStartPointX + 2.9, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 3.3, doubNextPortStartPointY + 0.2);
                                    arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                    arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].Text = Convert.ToString(intGlobalCableCounter) + " ББ";
                                    arrShapesBypass100MonFakeConnection[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 5.6, doubNextPortStartPointY, doubNextPortStartPointX + 5.6, doubNextPortStartPointY);
                                };
                            };
                            doubNextPortStartPointY -= 0.7;
                        };
                    };
                    //Конец рисования шасси байпасов
                    arrUplinkPortsOnBalancer[intCurrentBalancerChassis] = intCurrentPortInChassis;
                    //Console.WriteLine($"На последнем балансировщике {intCurrentBalancerChassis} последний порт {intCurrentPortInChassis}.");
                };

                // Крестуем IS40
                //bool boolBypassCross = false;
                bool boolBypassCross = true;

                if (boolBypassCross)
                    for (int intCurrentBalancer = 1; intCurrentBalancer <= intTotalBalancers; intCurrentBalancer++)
                    {
                        arrUplinkPortsOnBalancer[intCurrentBalancer] = 16;
                    }


                // Рисуем шасси IS40.
                if ((strBypassModel == "IS40" || strBypassModel == "MIX") && intTotalIs40Bypasses == 0)
                {
                    if (strBypassModel == "MIX")
                    {
                        intFirstBaypassChassis = 1;
                        intLastBypassChassis = intBypasses10New;
                    }
                    else
                    {
                        intFirstBaypassChassis = 1;
                        intLastBypassChassis = intTotalIs10Bypasses;
                    };
                    //Console.WriteLine($"Заполняем IS40.");
                    for (int intCurrentBypassDevice = intFirstBaypassChassis; intCurrentBypassDevice <= intLastBypassChassis; intCurrentBypassDevice++)
                    {
                        // Рисуем шасси.
                        listShapesBypass10Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 4.1));
                        strCurrentDeviceHostname = "IS40 (" + intCurrentBypassDevice + ")";
                        strNameForRackTables = $"{strFullSiteIndex}-IS40G-{intCurrentBypassDevice}";
                        listShapesBypass10Devices[listShapesBypass10Devices.Count - 1].Text = strCurrentDeviceHostname;
                        listShapesBypass10Devices[listShapesBypass10Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(219,112,147)";
                        // Рисуем порт Mgmt.
                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGNT ETH";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MGNT ETH";
                        // Рисуем линии для Mgmt.
                        listDeviceMgmtLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5));
                        listDeviceMgmtCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.7, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.1));
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                        listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5));
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "MGNT ETH";
                        // Сдвигаем координату Y вниз - для следующего IS40.
                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;
                        doubStartPointNextShapeY -= 8;

                        /////////////// Добавляем шасси в БД
                        intCurrentRackSlot++;
                        intDeviceId = FillRackSlot(command, strObjectIndex, intBypassRackId, intCurrentRackSlot, 50001, strNameForRackTables);

                        // MySQL Mgmt.
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'MGNT ETH', 1, 24);";
                        command.ExecuteNonQuery();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                        // Рисуем порты IS40 (10G).
                        intCurrentJournalItem = 0;
                        for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 3; intCurrentPortCounterInBypass++)
                        {
                            for (int intCurrentSubslotCounterInBypass = 1; intCurrentSubslotCounterInBypass <= 2; intCurrentSubslotCounterInBypass++)
                            {
                                intCurrentOverallLinkNumber++;
                                intCurrentBypassPort++;
                                // Net X/X/0
                                strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/0";
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Text = strCurrentPortName;
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                                // Соеденение двух Net-портов с LAN и WAN. Не соединяем Net-порты, к которым не приходят линки.
                                if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                                {
                                    //Odf Text
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 3.6, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 2.6, doubNextPortStartPointY - 0.3);
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].Text = arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Side_A_Port"];
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].LineStyle = "None";
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].Style = "None";
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                    arrShapesBypassNetFakeLine[intCurrentOverallLinkNumber, 0] = page1.DrawLine(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4);
                                    arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0] = page1.DrawOval(doubNextPortStartPointX - 1.1, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.7, doubNextPortStartPointY - 0.6);
                                    arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                    arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0].Text = intCurrentOverallLinkNumber + " L";
                                    arrShapesLanFakeCircles[intCurrentOverallLinkNumber].Text = intCurrentOverallLinkNumber + " L";
                                    arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4);
                                    //Соединительная линия.
                                    arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 0].AutoConnect(arrShapesLanFakeConnections[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                    //Дозаполнение ячейки массива словарей LAN-Bypass
                                    if (arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Comment"] == "Прямое включение") SideADevice = arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Side_A_Name"];
                                    else SideADevice = "ODF";
                                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Name", strCurrentDeviceHostname);
                                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Port", strCurrentPortName);
                                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Number", intCurrentOverallLinkNumber + " L");
                                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Name", $"{SideADevice} --- {strBypassModel} ({intCurrentBypassDevice})\n10GBASE-LR");
                                    //arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Name", $"ODF --- {strBypassModel} ({intCurrentBypassDevice})");
                                    list_CableJournal_LAN_Bypass.Add(new Dictionary<string, string>());
                                    foreach (string key in arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Keys)
                                    {
                                        list_CableJournal_LAN_Bypass[list_CableJournal_LAN_Bypass.Count - 1].Add(key, arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0][key]);
                                    };
                                    // MySQL
                                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                    command.ExecuteNonQuery();
                                    strNeighborPortId = arrShapesLanPorts[intCurrentOverallLinkNumber].Data1;
                                    strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                                    //Console.WriteLine($"porta: {strNeighborPortId}, portb: {strLocalPortId}");
                                    command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                                    command.ExecuteNonQuery();
                                };

                                // Net X/X/1
                                strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/1";
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Text = strCurrentPortName;
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                //Соеденение двух Net-портов с LAN и WAN. Не соединяем Net-порты, к которым не приходят линки.
                                if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                                {
                                    // Odf Text
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 3.6, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 2.6, doubNextPortStartPointY - 0.1);
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].Text = arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Side_A_Port"];
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].LineStyle = "None";
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].Style = "None";
                                    arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                    arrShapesBypassNetFakeLine[intCurrentOverallLinkNumber, 1] = page1.DrawLine(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.2);
                                    arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1] = page1.DrawOval(doubNextPortStartPointX - 1.6, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 1.2, doubNextPortStartPointY);
                                    arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                    arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1].Text = intCurrentOverallLinkNumber + " W";
                                    arrShapesWanFakeCircles[intCurrentOverallLinkNumber].Text = intCurrentOverallLinkNumber + " W";
                                    arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2);
                                    // Соединительная линия.
                                    arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 1].AutoConnect(arrShapesWanFakeConnections[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                    // Дозаполнение ячейки массива словарей WAN-Bypass
                                    if (arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Comment"] == "Прямое включение") SideADevice = arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Side_A_Name"];
                                    else SideADevice = "ODF";
                                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Name", strCurrentDeviceHostname);
                                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Port", strCurrentPortName);
                                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Number", intCurrentOverallLinkNumber + " W");
                                    //arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Name", $"ODF --- {strBypassModel} ({intCurrentBypassDevice})");
                                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Name", $"{SideADevice} --- {strBypassModel} ({intCurrentBypassDevice})\n10GBASE-LR");
                                    list_CableJournal_WAN_Bypass.Add(new Dictionary<string, string>());
                                    foreach (string key in arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Keys)
                                    {
                                        list_CableJournal_WAN_Bypass[list_CableJournal_WAN_Bypass.Count - 1].Add(key, arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1][key]);
                                    };
                                    // MySQL
                                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                    command.ExecuteNonQuery();
                                    strNeighborPortId = arrShapesWanPorts[intCurrentOverallLinkNumber].Data1;
                                    strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                                    command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                                    command.ExecuteNonQuery();
                                };

                                //-----------------------------------------------------------------------------
                                // Отслеживание лагов
                                arrLagIteration[intCurrentLagNumber]++;

                                //Console.WriteLine($"LAG: {intCurrentLagNumber}, Port: {arrLagIteration[intCurrentLagNumber]}");
                                if (arrLagIteration[intCurrentLagNumber] == 1)
                                {
                                    if (!boolEshelon && (32 - intCurrentPortInChassis < arrLanLagCounter[intCurrentLagNumber] / 2)) boolNotEnoughPortsForLag = true;

                                    //Console.WriteLine($"Первый порт LAG #{intCurrentLagNumber}, балансер: {intCurrentBalancerChassis + 1}, порт: {intCurrentPortInChassis + 1}, линков в лаге: {arrLanLagCounter[intCurrentLagNumber]}.");
                                    //Console.WriteLine($"21 - {intCurrentPortInChassis} < {arrLanLagCounter[intCurrentLagNumber] / 4}");
                                    //if (boolEshelon && ((21 - intCurrentPortInChassis) < arrLanLagCounter[intCurrentLagNumber] / 4)) Console.WriteLine($"21 - {intCurrentPortInChassis} < {arrLanLagCounter[intCurrentLagNumber]} / 4");
                                    if (boolEshelon && (20 - intCurrentPortInChassis < arrLanLagCounter[intCurrentLagNumber] / 4)) boolNotEnoughPortsForLag = true;
                                    //if (boolEshelon && (20 - intCurrentPortInChassis < arrLanLagCounter[intCurrentLagNumber] / 4)) Console.WriteLine("(+)");
                                }

                                if (boolEshelon)
                                {
                                    if (boolBypassCross)
                                    {
                                        intEshelonBalancerPortShift = 8;
                                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                        if ((intCurrentBypassPort - 1) % 4 == 0)
                                        {
                                            if (listHydraLines.Count == 0) intGlobalCableCounter++;
                                            if (intCurrentBalancerChassis == intTotalBalancers) intCurrentBalancerChassis = 1;
                                            else intCurrentBalancerChassis++;
                                            arrUplinkPortsOnBalancer[intCurrentBalancerChassis]++;
                                            intCurrentPortInChassis = arrUplinkPortsOnBalancer[intCurrentBalancerChassis];
                                        }
                                        else
                                        {
                                            //intCurrentPortInChassis = arrUplinkPortsOnBalancer[intCurrentBalancerChassis] + 1;
                                        }
                                        //else
                                        //{
                                        //    intCurrentPortInChassis = arrUplinkPortsOnBalancer[intCurrentBalancerChassis] + 8;
                                        //
                                        //}
                                        Console.WriteLine($"Порт байпаса: {intCurrentBypassPort}, порт хайвэя: {intCurrentPortInChassis}.");
                                        //Номер шасси меняется так же когда текущий порт - нечётный.
                                        //Но теперь для чётных линков текущий номер порта балансера = intCurrentBalancerChassis + 8

                                        Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                                        Console.WriteLine($"IS40 Chassis: {intCurrentBypassDevice}, IS40 Port: {intCurrentBypassPort}");
                                        Console.WriteLine($"ELB Chassis: {intCurrentBalancerChassis}, ELB Port: {arrUplinkPortsOnBalancer[intCurrentBalancerChassis]}");
                                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                                    }
                                    else
                                    {
                                        intEshelonBalancerPortShift = 8;
                                        if (listEshelonLanHydraLines.Count == 0) intGlobalCableCounter++;
                                        //if (((intCurrentBypassPort - 1) % 4 == 0 && boolFinishedFillingSingleBalancer) || iCurrentOverallMonPort == intTotalOverallLinkNumber * 2)
                                        Console.WriteLine($"Линк {iCurrentOverallMonPort} из {intTotalOverallLinkNumber * 2}");
                                        if (((intCurrentBypassPort - 1) % 4 == 0 && boolFinishedFillingSingleBalancer && iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2) || iCurrentOverallMonPort == intTotalOverallLinkNumber * 2)
                                        {
                                            intCurrentPortInChassis++;
                                            Console.WriteLine($"Порт байпаса: {intCurrentBypassPort}, порт хайвэя: {intCurrentPortInChassis}.");
                                        };
                                        if ((intCurrentBypassPort - 1) % 20 == 0)
                                        {
                                            if (intCurrentBalancerChassis > 0)
                                            {
                                                arrUplinkPortsOnBalancer[intCurrentBalancerChassis] = intCurrentPortInChassis - 1;
                                                Console.WriteLine($"На хайвэе {intCurrentBalancerChassis} последний заполненный порт: {intCurrentPortInChassis - 1}.");
                                            };
                                            boolNotEnoughPortsForLag = false;
                                            intCurrentBalancerChassis++;
                                            intCurrentPortInChassis = 17;
                                            boolFinishedFillingSingleBalancer = false;
                                        }    
                                    
                                    };
                                    if (arrLagIteration[intCurrentLagNumber] == arrLanLagCounter[intCurrentLagNumber])
                                    {
                                        intCurrentLagNumber++;
                                        arrLagIteration[intCurrentLagNumber] = 0;
                                    };
                                    Console.WriteLine($"Текущий присвоенный порт на шасси балансировщика: {intCurrentPortInChassis}");
                                }
                                else
                                {
                                    if (listHydraLines.Count == 0) intGlobalCableCounter++;
                                    if (!boolNoBalancer)                                    //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                                    {
                                        if (boolBypassCross)
                                        {

                                            if ((intCurrentBypassPort - 1) % 2 == 0)
                                            {
                                                if (intCurrentBalancerChassis == intTotalBalancers) intCurrentBalancerChassis = 1;
                                                else intCurrentBalancerChassis++;
                                                arrUplinkPortsOnBalancer[intCurrentBalancerChassis]++;
                                                intCurrentPortInChassis = arrUplinkPortsOnBalancer[intCurrentBalancerChassis];
                                            }

                                            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                                            Console.WriteLine($"IS40 Chassis {intCurrentBypassDevice}, IS40 Port: {intCurrentBypassPort}");
                                            Console.WriteLine($"ELB Chassis {intCurrentBalancerChassis}, ELB Port: {arrUplinkPortsOnBalancer[intCurrentBalancerChassis]}");
                                        }
                                        else
                                        {
                                            if ((intCurrentBypassPort - 1) % 2 == 0) intCurrentPortInChassis++;
                                            if (intCurrentPortInChassis == 33 || intCurrentBalancerChassis == 0 || boolNotEnoughPortsForLag)            // Правим тут!  LastPortOnBalancer
                                            //if (intCurrentPortInChassis == (LastPortOnBalancer + 1) || intCurrentBalancerChassis == 0 || boolNotEnoughPortsForLag)           
                                            {
                                                if (intCurrentBalancerChassis > 0)
                                                {
                                                    arrUplinkPortsOnBalancer[intCurrentBalancerChassis] = intCurrentPortInChassis - 1;
                                                    //Console.WriteLine($"На балансировщике {intCurrentBalancerChassis} последний заполненный порт: {intCurrentPortInChassis - 1}.");
                                                };
                                                boolNotEnoughPortsForLag = false;
                                                intCurrentBalancerChassis++;
                                                intCurrentPortInChassis = 17;
                                            };
                                            if (arrLagIteration[intCurrentLagNumber] == arrLanLagCounter[intCurrentLagNumber])
                                            {
                                                intCurrentLagNumber++;
                                                arrLagIteration[intCurrentLagNumber] = 0;
                                            };
                                        }

                                    };

                                    intEshelonBalancerPortShift = 0;
                                };

                                // Mon-порты (байпас - фильтр)
                                if (intFilterPortNoBalancer == intPortsNumberOnSingleFilter / 2)
                                {
                                    intCurrentFilterNoBalancer++;
                                    intFilterPortNoBalancer = 0;
                                };

                                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                //Mon/X/1
                                iCurrentOverallMonPort++;
                                strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/1";

                                if (!boolNoBalancer)
                                {
                                    if (boolEshelon)
                                    {
                                        strCableInHydra = " (AOC c" + (listEshelonWanHydraLines.Count + 1) + ")";
                                        intHydraEnd = listEshelonWanHydraLines.Count + 1;
                                    }
                                    else
                                    {
                                        if (intCurrentSubslotCounterInBypass == 1)
                                        {
                                            strCableInHydra = " (AOC c1)";
                                            intHydraEnd = 1;
                                        }
                                        else
                                        {
                                            strCableInHydra = " (AOC c3)";
                                            intHydraEnd = 3;
                                        };
                                    }
                                    Console.WriteLine($"Порт IS40: {strCurrentPortName}, Хайвэй: {intCurrentBalancerChassis}, Порт на хайвэе: {intCurrentPortInChassis}");
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass].Data1 = Convert.ToString(intCurrentBypassDevice);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass].Data2 = Convert.ToString(intCurrentBypassPort);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass].Text = strCurrentPortName;
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd] = new Dictionary<string, string>();
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Name", $"{strCurrentDeviceHostname} --- ELB-0133 ({intCurrentBalancerChassis})");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + strCableInHydra);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Device_A_Name", strCurrentDeviceHostname);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Type", "NR-QSFP-4X10G-AOC 40G QSFP");

                                    //////////  MySQL
                                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                    command.ExecuteNonQuery();
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));
                                }
                                else if (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                                {
                                    intFilterPortNoBalancer++;
                                    strCableInHydra = "";
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Text = strCurrentPortName;
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                    //Draw Line And Fake Rectangle
                                    arrShapesBypass100MonFakeLine[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1] = page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 3.7, doubNextPortStartPointY - 0.2);
                                    arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1] = page1.DrawOval(doubNextPortStartPointX + 3.2, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 3.6, doubNextPortStartPointY);
                                    arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                    arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1].Text = intFilterPortNoBalancer * 2 - 1 + " БФ";
                                    arrShapesBypass100MonFakeConnection[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1] = page1.DrawRectangle(doubNextPortStartPointX + 3.7, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 3.7, doubNextPortStartPointY - 0.2);
                                    intGlobalCableCounter++;
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0] = new Dictionary<string, string>();
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Number", intFilterPortNoBalancer * 2 - 1 + " БФ");
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Name", strCurrentDeviceHostname + " --- " + strFilterModel + " (1)");
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Type", "NR-SFP-10G-AOC 10G SFP+ <=> SFP+");

                                    //////////  MySQL
                                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                    command.ExecuteNonQuery();
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));
                                };
                                // Mon/X/0
                                iCurrentOverallMonPort++;
                                strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/0";

                                if (!boolNoBalancer)
                                {
                                    if (boolEshelon)
                                    {
                                        if (boolBypassCross)
                                        {
                                            strCableInHydra = " (AOC c" + (listEshelonWanHydraLines.Count + 1) + ")";
                                            intHydraEnd = listEshelonWanHydraLines.Count + 1;
                                        }
                                        else
                                        {
                                            strCableInHydra = " (AOC c" + (listEshelonWanHydraLines.Count + 1) + ")";
                                            intHydraEnd = listEshelonWanHydraLines.Count + 1;
                                        }
                                    }
                                    else
                                    {
                                        if (intCurrentSubslotCounterInBypass == 1)
                                        {
                                            strCableInHydra = " (AOC c2)";
                                            intHydraEnd = 2;
                                        }
                                        else
                                        {
                                            strCableInHydra = " (AOC c4)";
                                            intHydraEnd = 4;
                                        };
                                    };
                                    Console.WriteLine($"Порт IS40: {strCurrentPortName}, Хайвэй: {intCurrentBalancerChassis}, Порт на хайвэе: {intCurrentPortInChassis + intEshelonBalancerPortShift}");
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intCurrentSubslotCounterInBypass * 2] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intCurrentSubslotCounterInBypass * 2].Data1 = Convert.ToString(intCurrentBypassDevice);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intCurrentSubslotCounterInBypass * 2].Data2 = Convert.ToString(intCurrentBypassPort);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intCurrentSubslotCounterInBypass * 2].Text = strCurrentPortName;
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intCurrentSubslotCounterInBypass * 2].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intHydraEnd] = new Dictionary<string, string>();
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intHydraEnd].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intHydraEnd].Add("Cable_Name", $"{strCurrentDeviceHostname} --- ELB-0133 ({intCurrentBalancerChassis})");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intHydraEnd].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + strCableInHydra);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intHydraEnd].Add("Device_A_Name", strCurrentDeviceHostname);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intHydraEnd].Add("Cable_Type", "NR-QSFP-4X10G-AOC 40G QSFP");

                                    // MySQL
                                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                    command.ExecuteNonQuery();
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis + intEshelonBalancerPortShift, intHydraEnd].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));
                                }
                                else if (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                                {
                                    strCableInHydra = "";
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Text = strCurrentPortName;
                                    arrShapesBypass10_MonPorts[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                    //Draw Line And Fake Rectangle
                                    arrShapesBypass100MonFakeLine[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2] = page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 3.3, doubNextPortStartPointY - 0.4);
                                    arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2] = page1.DrawOval(doubNextPortStartPointX + 2.8, doubNextPortStartPointY - 0.6, doubNextPortStartPointX + 3.2, doubNextPortStartPointY - 0.2);
                                    arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                    arrShapesBypassMonFakeCircles[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2].Text = intFilterPortNoBalancer * 2 + " БФ";
                                    arrShapesBypass100MonFakeConnection[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2] = page1.DrawRectangle(doubNextPortStartPointX + 3.3, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 3.3, doubNextPortStartPointY - 0.4);
                                    intGlobalCableCounter++;
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0] = new Dictionary<string, string>();
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Cable_Number", intFilterPortNoBalancer * 2 + " БФ");
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Cable_Name", strCurrentDeviceHostname + " --- " + strFilterModel + " (1)");
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Cable_Type", "NR-SFP-10G-AOC 10G SFP+ <=> SFP+");

                                    // MySQL
                                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                    command.ExecuteNonQuery();
                                    arr_CableJournal_Bypass_Balancer[intCurrentFilterNoBalancer, intFilterPortNoBalancer * 2, 0].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));

                                };

                                if (!boolNoBalancer && (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)) boolFinishedFillingSingleBalancer = true;

                                ///////////////     Draw Hydra Connector    //////////////////////////
                                if (!boolNoBalancer && !boolEshelon && (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2))
                                {
                                    listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.4));
                                    listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.2));
                                    if (listHydraLines.Count == 4 || iCurrentOverallMonPort == intTotalOverallLinkNumber * 2)
                                    {
                                        /*
                                        if (boolBypassCross)
                                        {
                                            if (intCurrentBalancerChassis == intTotalBalancers) intCurrentBalancerChassis = 1;
                                            else intCurrentBalancerChassis++;
                                            arrUplinkPortsOnBalancer[intCurrentBalancerChassis]++;
                                            intCurrentPortInChassis = arrUplinkPortsOnBalancer[intCurrentBalancerChassis];
                                            //Console.WriteLine($"ELB Chassis {intCurrentBalancerChassis}, ELB Port: {arrUplinkPortsOnBalancer[intCurrentBalancerChassis]}");
                                        }
                                        */

                                        list_Specification.Add("duplex LC/UPC-LC/UPC, SM");
                                        list_Specification.Add("NR-QSFP-4X10G-AOC 40G QSFP");
                                        if (iCurrentOverallMonPort == intTotalOverallLinkNumber * 2 && listHydraLines.Count < 4) doubShiftY = 0.7;
                                        else doubShiftY = 0;
                                        listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.4 - doubShiftY, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY + 0.5 - doubShiftY));
                                        vsoWindow.DeselectAll();
                                        foreach (Visio.Shape objHydraSingleLine in listHydraLines)
                                        {
                                            vsoWindow.Select(objHydraSingleLine, 2);
                                        };
                                        vsoSelection = vsoWindow.Selection;
                                        arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis] = vsoSelection.Group();
                                        arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(intCurrentBypassDevice);
                                        arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis].Data3 = Convert.ToString(intCurrentBypassDevice);
                                        intBypassIs40CurrentHydra++;
                                        listHydraLines.Clear();

                                        // Draw Line And Fake Rectangle
                                        arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawLine(doubNextPortStartPointX + 2.6, doubNextPortStartPointY - doubShiftY, doubNextPortStartPointX + 5.6, doubNextPortStartPointY - doubShiftY);
                                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawOval(doubNextPortStartPointX + 2.9, doubNextPortStartPointY - 0.2 - doubShiftY, doubNextPortStartPointX + 3.3, doubNextPortStartPointY + 0.2 - doubShiftY);
                                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].Text = Convert.ToString(intGlobalCableCounter) + " ББ";
                                        arrShapesBypass100MonFakeConnection[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 5.6, doubNextPortStartPointY - doubShiftY, doubNextPortStartPointX + 5.6, doubNextPortStartPointY - doubShiftY);
                                    };
                                };

                                if (boolEshelon)
                                {
                                    // Если сейчас будет нарисован первый луч гидры, фиксируем координаты Y для WAN и LAN лучей
                                    if (listEshelonWanHydraLines.Count == 0)
                                    {
                                        doubLongWanHydraStartCoordinateY = doubNextPortStartPointY - 0.2;
                                        doubLongLanHydraStartCoordinateY = doubNextPortStartPointY - 0.4;
                                    };
                                    //if (listEshelonWanHydraLines.Count <= 2) noMoreTwoLinksInHydra = true;
                                    // Рисуем два луча линка - WAN и LAN
                                    //listEshelonWanHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 2.7, doubNextPortStartPointY - 0.2));
                                    //listEshelonLanHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 2.9, doubNextPortStartPointY - 0.4));

                                    if (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                                    {
                                        listEshelonWanHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 2.7, doubNextPortStartPointY - 0.2));
                                        listEshelonLanHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 2.9, doubNextPortStartPointY - 0.4));
                                    };
                                    /*
                                    */

                                    // Если только что был нарисован последний 4й луч гидры, рисуем планки для LAN и WAN гидр и группируем хвосты гидры в общую фигуру
                                    // Либо если мы рисуем последний линк.
                                    if (listEshelonWanHydraLines.Count == 4 || iCurrentOverallMonPort == intTotalOverallLinkNumber * 2)
                                    {
                                        if (iCurrentOverallMonPort == intTotalOverallLinkNumber * 2 && listEshelonWanHydraLines.Count == 1) doubNextPortStartPointY -= 0.6;
                                        // Группирование гидры WAN
                                        listEshelonWanHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.7, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 2.7, doubLongWanHydraStartCoordinateY));
                                        vsoWindow.DeselectAll();
                                        foreach (Visio.Shape objHydraSingleLine in listEshelonWanHydraLines)
                                        {
                                            objHydraSingleLine.CellsU["LineColor"].FormulaForceU = "THEMEGUARD(RGB(0,0,255))";
                                            vsoWindow.Select(objHydraSingleLine, 2);
                                        };
                                        vsoSelection = vsoWindow.Selection;
                                        //Console.WriteLine($"");
                                        arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis] = vsoSelection.Group();
                                        arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(intCurrentBypassDevice);
                                        arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis].Data3 = Convert.ToString(intCurrentBypassDevice);
                                        intBypassIs40CurrentHydra++;
                                        strHydraName = Convert.ToString(intGlobalCableCounter * 2 - 1) + " ББ";

                                        // Draw Line And Fake Rectangle
                                        arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawLine(doubNextPortStartPointX + 2.7, doubNextPortStartPointY + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2, doubNextPortStartPointX + 5.6, doubNextPortStartPointY + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2);
                                        arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis].CellsU["LineColor"].FormulaForceU = "THEMEGUARD(RGB(0,0,255))";
                                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawOval(doubNextPortStartPointX + 3.1, doubNextPortStartPointY - 0.2 + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2, doubNextPortStartPointX + 3.5, doubNextPortStartPointY + 0.2 + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2);
                                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].Text = strHydraName;
                                        arrShapesBypass100MonFakeConnection[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 5.6, doubNextPortStartPointY + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2, doubNextPortStartPointX + 5.6, doubNextPortStartPointY + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2);
                                        //Console.WriteLine($"Линк: {strHydraName}, Балансировщик: {intCurrentBalancerChassis}, Порт: {intCurrentPortInChassis}");

                                        // Группирование гидры LAN
                                        listEshelonLanHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.9, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 2.9, doubLongLanHydraStartCoordinateY));
                                        vsoWindow.DeselectAll();
                                        foreach (Visio.Shape objHydraSingleLine in listEshelonLanHydraLines)
                                        {
                                            objHydraSingleLine.CellsU["LineColor"].FormulaForceU = "THEMEGUARD(RGB(255,102,0))";
                                            vsoWindow.Select(objHydraSingleLine, 2);
                                        };
                                        vsoSelection = vsoWindow.Selection;
                                        arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis + 8] = vsoSelection.Group();
                                        arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis + 8].Data1 = Convert.ToString(intCurrentBypassDevice);
                                        arrBypassIs40HydraConnectors[intCurrentBalancerChassis, intCurrentPortInChassis + 8].Data3 = Convert.ToString(intCurrentBypassDevice);
                                        intBypassIs40CurrentHydra++;
                                        strHydraName = Convert.ToString(intGlobalCableCounter * 2) + " ББ";
                                        //Draw Line And Fake Rectangle
                                        arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis + 8] = page1.DrawLine(doubNextPortStartPointX + 2.9, doubNextPortStartPointY - 0.4 + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2, doubNextPortStartPointX + 5.6, doubNextPortStartPointY - 0.4 + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2);
                                        arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis + 8].CellsU["LineColor"].FormulaForceU = "THEMEGUARD(RGB(255,102,0))";
                                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis + 8] = page1.DrawOval(doubNextPortStartPointX + 3.5, doubNextPortStartPointY - 0.6 + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2, doubNextPortStartPointX + 3.9, doubNextPortStartPointY - 0.2 + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2);
                                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis + 8].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                        arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis + 8].Text = strHydraName;
                                        arrShapesBypass100MonFakeConnection[intCurrentBalancerChassis, intCurrentPortInChassis + 8] = page1.DrawRectangle(doubNextPortStartPointX + 5.6, doubNextPortStartPointY - 0.4 + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2, doubNextPortStartPointX + 5.6, doubNextPortStartPointY - 0.4 + (doubLongLanHydraStartCoordinateY - doubNextPortStartPointY) / 2);
                                        //Console.WriteLine($"Линк: {strHydraName}, Балансировщик: {intCurrentBalancerChassis}, Порт: {intCurrentPortInChassis}");
                                        vsoWindow.DeselectAll();
                                        listEshelonWanHydraLines.Clear();
                                        listEshelonLanHydraLines.Clear();

                                        if (iCurrentOverallMonPort == intTotalOverallLinkNumber * 2 && listEshelonWanHydraLines.Count == 1) doubNextPortStartPointY += 0.6;
                                    };

                                };


                                //Сдвиг вниз следующей фигуры байпаса
                                doubBypassEndX = doubStartPointNextShapeX + 1;
                                doubNextPortStartPointY -= 0.7;                 // Move down to next port

                                //Группируем шасси IS40 с его портами

                                if (intCurrentPortCounterInBypass == 4)
                                {
                                    //Console.WriteLine("Net-Линк #4");
                                    vsoWindow.DeselectAll();
                                    vsoWindow.Select(listShapesIbs1UpDevices[listShapesIbs1UpDevices.Count - 1], 2);
                                    vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1], 2);
                                    vsoWindow.Select(listDeviceMgmtLines[listDeviceMgmtLines.Count - 1], 2);
                                    vsoWindow.Select(listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1], 2);
                                    vsoWindow.Select(listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1], 2);
                                    for (int intCurrentLinkOnBypass = intCurrentOverallLinkNumber - 3; intCurrentLinkOnBypass <= intCurrentOverallLinkNumber; intCurrentLinkOnBypass++)
                                    {
                                        vsoWindow.Select(arrShapesBypass100NetPorts[intCurrentLinkOnBypass, 0], 2);
                                        vsoWindow.Select(arrShapesBypassNetFakeLine[intCurrentLinkOnBypass, 0], 2);
                                        vsoWindow.Select(arrShapesBypassNetFakeCircles[intCurrentLinkOnBypass, 0], 2);
                                        vsoWindow.Select(arrShapesBypassNetFakeConnection[intCurrentLinkOnBypass, 0], 2);
                                        vsoWindow.Select(arrShapesBypass10_Odf[intCurrentLinkOnBypass, 0], 2);
                                        vsoWindow.Select(arrShapesBypass100NetPorts[intCurrentLinkOnBypass, 1], 2);
                                        vsoWindow.Select(arrShapesBypassNetFakeLine[intCurrentLinkOnBypass, 1], 2);
                                        vsoWindow.Select(arrShapesBypassNetFakeCircles[intCurrentLinkOnBypass, 1], 2);
                                        vsoWindow.Select(arrShapesBypassNetFakeConnection[intCurrentLinkOnBypass, 1], 2);
                                        vsoWindow.Select(arrShapesBypass10_Odf[intCurrentLinkOnBypass, 1], 2);
                                    };

                                    vsoWindow.Select(arrShapesBypass100MonPorts[int100mDevice1, int100mPort1], 2);
                                    vsoWindow.Select(arrShapesBypass100MonPorts[int100mDevice1, int100mPort1 + 1], 2);
                                    vsoWindow.Select(arrShapesBypass100MonFakeLine[int100mDevice1, int100mPort1], 2);
                                    vsoWindow.Select(arrShapesBypass100MonFakeLine[int100mDevice1, int100mPort1 + 1], 2);
                                    vsoWindow.Select(arrShapesBypassMonFakeCircles[int100mDevice1, int100mPort1], 2);
                                    vsoWindow.Select(arrShapesBypassMonFakeCircles[int100mDevice1, int100mPort1 + 1], 2);
                                    vsoWindow.Select(arrShapesBypass100MonFakeConnection[int100mDevice1, int100mPort1], 2);
                                    vsoWindow.Select(arrShapesBypass100MonFakeConnection[int100mDevice1, int100mPort1 + 1], 2);
                                    vsoWindow.Select(arrShapesBypass10_MonPorts[int100mDevice2, int100mPort2, 0], 2);
                                    vsoWindow.Select(arrShapesBypass10_MonPorts[int100mDevice2, int100mPort2 + 1, 0], 2);
                                    vsoWindow.Select(arrShapesBypass100MonFakeLine[int100mDevice2, int100mPort2], 2);
                                    vsoWindow.Select(arrShapesBypass100MonFakeLine[int100mDevice2, int100mPort2 + 1], 2);
                                    vsoWindow.Select(arrShapesBypassMonFakeCircles[int100mDevice2, int100mPort2], 2);
                                    vsoWindow.Select(arrShapesBypassMonFakeCircles[int100mDevice2, int100mPort2 + 1], 2);
                                    vsoWindow.Select(arrShapesBypass100MonFakeConnection[int100mDevice2, int100mPort2], 2);
                                    vsoWindow.Select(arrShapesBypass100MonFakeConnection[int100mDevice2, int100mPort2 + 1], 2);
                                    /*
                                    */
                                    vsoSelection = vsoWindow.Selection;
                                    vsoSelection.Group();

                                };

                            };

                        };
                    };  //Конец рисования шасси байпасов
                    if (!boolBypassCross) arrUplinkPortsOnBalancer[intCurrentBalancerChassis] = intCurrentPortInChassis;
                    Console.WriteLine($"Записали в шасси балансера {intCurrentBalancerChassis} последний номер порта: {arrUplinkPortsOnBalancer[intCurrentBalancerChassis]}");
                    if (boolEshelon) arrUplinkPortsOnBalancer[intCurrentBalancerChassis]--;                                                                               //Правим тут!
                    //if (boolEshelon && intCurrentBalancerChassis < intTotalBalancers && intTotalBalancers > 1) arrUplinkPortsOnBalancer[intCurrentBalancerChassis]--;
                    //if (boolEshelon && intCurrentBalancerChassis < Convert.ToInt32(strBalancerNumberFromInput)) arrUplinkPortsOnBalancer[intCurrentBalancerChassis]--;

                    //Console.WriteLine($"На последнем балансировщике {intCurrentBalancerChassis} последний порт {intCurrentPortInChassis}.");
                };


                intBypassIs40HydrasTotal = intBypassIs40CurrentHydra;

                // Рисуем IS40 (40G).
                List<Visio.Shape> listShapesBypass40Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass40NetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass40MonPorts = new List<Visio.Shape>();
                intCurrentBypassPort = 0;
                iCurrentOverallMonPort = 0;
                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs40Bypasses; intCurrentBypassDevice++)
                {
                    listShapesBypass40Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 2));
                    strCurrentDeviceHostname = "IS40 (" + intCurrentBypassDevice + ")";
                    strNameForRackTables = $"{strFullSiteIndex}-IS40G-{intCurrentBypassDevice}";
                    listShapesBypass40Devices[listShapesBypass40Devices.Count - 1].Text = strCurrentDeviceHostname;
                    listShapesBypass40Devices[listShapesBypass40Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(255,165,0)";

                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGNT ETH";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();

                    // Рисуем линии для Mgmt.
                    listDeviceMgmtLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5));
                    listDeviceMgmtCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.7, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.1));
                    listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                    listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                    listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5));
                    listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "MGNT ETH";

                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;
                    doubStartPointNextShapeY -= 4.5;

                    /////////////// Добавляем шасси в БД
                    intCurrentRackSlot++;
                    intDeviceId = FillRackSlot(command, strObjectIndex, intBypassRackId, intCurrentRackSlot, 50001, strNameForRackTables);

                    // MySQL Mgmt.
                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'MGNT ETH', 1, 24);";
                    command.ExecuteNonQuery();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                    // Рисуем IS40 (40G).
                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 3; intCurrentPortCounterInBypass++)
                    {
                        // Отрисовка портов Net0
                        intCurrentBypassPort++;
                        intCurrentOverallLinkNumber++;
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/0";
                        arrShapesBypass40_NetPorts[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                        arrShapesBypass40_NetPorts[intCurrentOverallLinkNumber, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass40_NetPorts[intCurrentOverallLinkNumber, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass40_NetPorts[intCurrentOverallLinkNumber, 0].Text = strCurrentPortName;
                        arrShapesBypass40_NetPorts[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                        {
                            //Odf Text
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 3.6, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 2.6, doubNextPortStartPointY - 0.3);
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].Text = arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Side_A_Port"];
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].LineStyle = "None";
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].Style = "None";
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            //Fakes
                            arrShapesBypassNetFakeLine[intCurrentOverallLinkNumber, 0] = page1.DrawLine(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4);
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0] = page1.DrawOval(doubNextPortStartPointX - 1.1, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.7, doubNextPortStartPointY - 0.6);
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 0].Text = intCurrentOverallLinkNumber + " L";
                            arrShapesLanFakeCircles[intCurrentOverallLinkNumber].Text = intCurrentOverallLinkNumber + " L";
                            arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.4);
                            //Connect
                            arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 0].AutoConnect(arrShapesLanFakeConnections[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            //Дозаполнение ячейки массива словарей LAN-Bypass
                            if (arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Comment"] == "Прямое включение") SideADevice = arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0]["Side_A_Name"];
                            else SideADevice = "ODF";
                            arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Name", strCurrentDeviceHostname);
                            arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Port", strCurrentPortName);
                            arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Number", intCurrentOverallLinkNumber + " L");
                            arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Name", $"{SideADevice} --- {strBypassModel} ({intCurrentBypassDevice})\n40GBASE-LR4");
                            //arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Name", $"ODF --- {strBypassModel} ({intCurrentBypassDevice})");
                            list_CableJournal_LAN_Bypass.Add(new Dictionary<string, string>());

                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 10, 1664);";
                            command.ExecuteNonQuery();
                            strNeighborPortId = arrShapesLanPorts[intCurrentOverallLinkNumber].Data1;
                            strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                            //Console.WriteLine($"porta: {arrShapesWanPorts[intCurrentOverallLinkNumber].Data1}");
                            command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                            command.ExecuteNonQuery();

                            foreach (string key in arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Keys)
                            {
                                list_CableJournal_LAN_Bypass[list_CableJournal_LAN_Bypass.Count - 1].Add(key, arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0][key]);
                            };
                        };

                        // Отрисовка портов Net1
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/1";
                        arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1);
                        arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Text = strCurrentPortName;
                        arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                        {
                            //Odf Text
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 3.6, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 2.6, doubNextPortStartPointY - 0.1);
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].Text = arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Side_A_Port"];
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].LineStyle = "None";
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].Style = "None";
                            arrShapesBypass10_Odf[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            //Fakes
                            arrShapesBypassNetFakeLine[intCurrentOverallLinkNumber, 1] = page1.DrawLine(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.2);
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1] = page1.DrawOval(doubNextPortStartPointX - 1.6, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 1.2, doubNextPortStartPointY);
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrShapesBypassNetFakeCircles[intCurrentOverallLinkNumber, 1].Text = intCurrentOverallLinkNumber + " W";
                            arrShapesWanFakeCircles[intCurrentOverallLinkNumber].Text = intCurrentOverallLinkNumber + " W";
                            arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 5.5, doubNextPortStartPointY - 0.2);
                            //Connect
                            arrShapesBypassNetFakeConnection[intCurrentOverallLinkNumber, 1].AutoConnect(arrShapesWanFakeConnections[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            //Дозаполнение ячейки массива словарей WAN-Bypass
                            if (arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Comment"] == "Прямое включение") SideADevice = arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1]["Side_A_Name"];
                            else SideADevice = "ODF";
                            arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Name", strCurrentDeviceHostname);
                            arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Port", strCurrentPortName);
                            arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Number", intCurrentOverallLinkNumber + " W");
                            //arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Name", $"ODF --- {strBypassModel} ({intCurrentBypassDevice})");
                            arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Name", $"{SideADevice} --- {strBypassModel} ({intCurrentBypassDevice})\n40GBASE-LR4");
                            list_CableJournal_WAN_Bypass.Add(new Dictionary<string, string>());

                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 10, 1664);";
                            command.ExecuteNonQuery();
                            strNeighborPortId = arrShapesWanPorts[intCurrentOverallLinkNumber].Data1;
                            strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                            command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                            command.ExecuteNonQuery();

                            foreach (string key in arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Keys)
                            {
                                list_CableJournal_WAN_Bypass[list_CableJournal_WAN_Bypass.Count - 1].Add(key, arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1][key]);
                            };
                        };

                        // Отрисовка портов Mon
                        //Раундробин (формула).
                        intCurrentBalancerChassis++;
                        //WAN-линки - в нечётные порты балансировщика (1, 3, 5, 7,...)
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/1";
                        intCurrentPortInChassis = arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 1;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Text = strCurrentPortName;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                        {
                            //КЖ линков WAN
                            intGlobalCableCounter++;
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0] = new Dictionary<string, string>();
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + " ББ");
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Name", $"{strCurrentDeviceHostname} --- ELB-0133 ({intCurrentBalancerChassis})");
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Type", "NR-QSFP-40G-AOC 40G QSFP <=> QSFP");

                            //Draw Line And Fake Rectangle
                            arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.2);
                            arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawOval(doubNextPortStartPointX + 2.9, doubNextPortStartPointY, doubNextPortStartPointX + 3.3, doubNextPortStartPointY - 0.4);
                            arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].Text = Convert.ToString(intGlobalCableCounter) + " ББ";
                            arrShapesBypass100MonFakeConnection[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.2);

                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 15, 1672);";
                            command.ExecuteNonQuery();
                            arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(LastUsedId(command, "Port"));
                        };
                        //LAN-линки - в чётные порты балансировщика (2, 4, 6, 8,...)
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                        intCurrentPortInChassis = arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 2;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Text = strCurrentPortName;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                        {
                            //КЖ линков LAN
                            intGlobalCableCounter++;
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0] = new Dictionary<string, string>();
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + " ББ");
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Name", $"{strCurrentDeviceHostname} --- ELB-0133 ({intCurrentBalancerChassis})");
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Type", "NR-QSFP-40G-AOC 40G QSFP <=> QSFP");

                            //Draw Line And Fake Rectangle
                            arrShapesBypass100MonFakeLine[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.4);
                            arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawOval(doubNextPortStartPointX + 3.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 3.9, doubNextPortStartPointY - 0.6);
                            arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrShapesBypassMonFakeCircles[intCurrentBalancerChassis, intCurrentPortInChassis].Text = Convert.ToString(intGlobalCableCounter) + " ББ";
                            arrShapesBypass100MonFakeConnection[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 5.5, doubNextPortStartPointY - 0.4);

                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 15, 1672);";
                            command.ExecuteNonQuery();
                            arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(LastUsedId(command, "Port"));
                        };


                        //После отрисовки пары WAN-LAN сдвиг указателя на 2 порта.
                        arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] += 2;

                        //После полного прогона балансировщиков указатель перемещается в начало - к первому балансировщику.
                        if (intCurrentBalancerChassis == intTotalBalancers) intCurrentBalancerChassis = 0;

                        //Сдвиг вниз следующей фигуры байпаса
                        doubNextPortStartPointY -= 0.7;
                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    };

                }




                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                doubLastBypassPortY = doubStartPointNextShapeY + 6;
                doubBypassEndX = doubStartPointNextShapeX + 1;
                doubBypassEndX = doubStartPointNextShapeX + 1;
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                int intJournalCurrentRow = 0;
                int intLastUplinkPortOnBalancer = 0;
                int intUsedUplinkHydrasCounter = 0;
                int intLastDownlinkPortOnBalancer = 0;
                int intUsedDownlinkHydrasCounter = 0;
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                //ODF-планка
                Visio.Shape shapeOdfBar = page1.DrawRectangle(doubFirstBypassPortX, doubFirstBypassPortY, doubFirstBypassPortX + 0.6, doubLastBypassPortY - 1);
                shapeOdfBar.Text = "ODF";
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //Console.WriteLine($"Отрисовка байпасов завершена.");
                intGlobalCableCounter = 0;
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw Balancers   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                List<Visio.Shape> listShapesBalancerDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBalancerUplinkPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBalancerDownlinkPorts = new List<Visio.Shape>();

                Visio.Shape[,] arrShapesBalancer100UplinkPorts = new Visio.Shape[10, 100];
                Visio.Shape[,] arrShapesBalancer1040DownlinkPorts = new Visio.Shape[50, 100];

                Visio.Shape[,] arrShapesBalancerFakeUplinkLines = new Visio.Shape[20, 100];
                Visio.Shape[,] arrShapesBalancerFakeUplinkConnections = new Visio.Shape[20, 100];
                Visio.Shape[,] arrShapesBalancerFakeUplinkCircles = new Visio.Shape[20, 100];

                Visio.Shape[,] arrShapesBalancerFakeDownlinkLines = new Visio.Shape[100, 100];
                Visio.Shape[,] arrShapesBalancerFakeDownlinkConnections = new Visio.Shape[100, 100];
                Visio.Shape[,] arrShapesBalancerFakeDownlinkCircles = new Visio.Shape[100, 100];

                Visio.Shape[,] arrShapesBalancerPeremychkaPorts = new Visio.Shape[100, 100];
                Visio.Shape[,] arrShapesBalancerPeremychkaLines = new Visio.Shape[100, 100];
                Visio.Shape[,] arrShapesBalancerFakePeremychkaConnections = new Visio.Shape[100, 100];
                Visio.Shape[,] arrShapesBalancerFakePeremychkaCircles = new Visio.Shape[100, 100];

                Visio.Shape[] arrShapesBalancerDevices = new Visio.Shape[10];
                Visio.Shape[] arrShapesFilterDevices = new Visio.Shape[100];

                doubStartPointNextShapeX += 20;

                intdoubTopLineY = doubUpperStartPoint + 2;

                doubNextPortStartPointX = doubStartPointNextShapeX;
                doubNextPortStartPointY = doubStartPointNextShapeY;

                int intGapBwBalancers = 15;
                int intFilterPointerCross;
                int intFilterPointerStraight = 0;
                int intHydraStartPoint;

                int intHydraPointerCross;
                int intHydraPointerStraight = 0;
                int intCurrentHighwayPeremyckaNumber = 0;
                int intCurrentHighwayCurrentPeremychkaPort = 0;
                int intOverallPeremychkaNumber = 0;

                //Вычисление общего количеств портов 10G к балансировщикам. Удвоенное количество операторских 10G-линков. Возможно, округлить в большую сторону
                int intTotalBalancerUplink100Ports = intLinkCounter100 * 2 + intLinkCounter40 * 2;
                int intBalancerUplinkPortPerDevice = 0;
                int intTotalFiltersQuantity = Convert.ToInt32(strFilterNumberFromInput);
                int intTotalBalancersQuantity = 0;
                int intHydrasQuantityBwOnePair = 0;
                int intSubseqHydrasCounter;
                int intCrossPortsOnEachBalancer = 0;

                int intBalancerRecalculatedPort;
                int intMaximumUplinkPortsOnBalancer;
                int intMaximBalancerPortNumber;

                if (boolEshelon)
                {
                    intMaximumUplinkPortsOnBalancer = 10;
                    intMaximBalancerPortNumber = 26;
                }
                else
                {
                    intMaximumUplinkPortsOnBalancer = 16;
                    intMaximBalancerPortNumber = 32;
                };

                int intTotalFilterHydrasQuantity = intHydrasOnFilter * intTotalFiltersQuantity; //strFilterNumberFromInput;

                if (!boolNoBalancer)
                {
                    //Console.WriteLine($"Отрисовка балансировщиков начата.");
                    intTotalBalancersQuantity = Convert.ToInt32(strBalancerNumberFromInput);
                    intCrossPortsOnEachBalancer = CalculateDevicesQuantity(intTotalBalancerUplink100Ports, intTotalBalancersQuantity);
                    if (intCrossPortsOnEachBalancer % 2 > 0) intCrossPortsOnEachBalancer++;
                    //Console.WriteLine($"По {intCrossPortsOnEachBalancer} портов на каждом хайвэе.");
                    intBalancerUplinkPortPerDevice = intTotalFilterHydrasQuantity / intTotalBalancersQuantity;
                    //if (intTotalBalancerUplink100Ports % intTotalBalancersQuantity > 0) Console.WriteLine("Нацело не делится");
                    intHydrasQuantityBwOnePair = CalculateDevicesQuantity(intTotalFilterHydrasQuantity, intTotalBalancersQuantity * intTotalFiltersQuantity);
                };
                //Console.WriteLine($"Балансировщики: {intTotalBalancersQuantity}, Фильтры: {intTotalFiltersQuantity}.");
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw ELB Device ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                intCurrentRackSlot = 10;

                doubStartPointNextShapeX += intTotalIs100Bypasses * 2 + intTotalIs40Bypasses + intTotalIs10Bypasses * 0.5 + 3;
                strCurrentDeviceHostname = "";
                //Для варианта с балансировщиком. Коммутация линков 10G гидрами.
                if (!boolNoBalancer)
                {
                    //Console.WriteLine($"Балансировщик !!!");
                    for (int intCurrentBalancerFrame = 1; intCurrentBalancerFrame <= intTotalBalancersQuantity; intCurrentBalancerFrame++)
                    {
                        arrShapesBalancerDevices[intCurrentBalancerFrame] = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 9);
                        strCurrentDeviceHostname = "ELB-0133 (" + intCurrentBalancerFrame + ")";
                        strNameForRackTables = $"{strFullSiteIndex}-ELB-{intCurrentBalancerFrame}";
                        arrShapesBalancerDevices[intCurrentBalancerFrame].Text = strCurrentDeviceHostname;
                        arrShapesBalancerDevices[intCurrentBalancerFrame].get_Cells("FillForegnd").FormulaU = "=RGB(143,188,143)";
                        arrShapesBalancerDevices[intCurrentBalancerFrame].Data3 = strCurrentDeviceHostname;

                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGMT";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MGMT";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(intCurrentBalancerFrame);

                        //Mgmt Lines
                        listDeviceMgmtLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5));
                        listDeviceMgmtCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.7, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.1));
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                        listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.5));
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "MGMT";
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data3 = Convert.ToString(intCurrentBalancerFrame);


                        intStartBalancerPort = 17;

                        /////////////// Добавляем шасси в БД
                        intCurrentRackSlot++;
                        intDeviceId = FillRackSlot(command, strObjectIndex, intBalancerRackId, intCurrentRackSlot, 50000, strNameForRackTables);

                        // MySQL Mgmt.
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'MGMT', 1, 24);";
                        command.ExecuteNonQuery();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw ELB to 100G Uplink Ports (IS100)   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        if (intTotalBalancerUplink100Ports > 0)
                        {
                            for (int intCurrentBalancerUplinkPort = intStartBalancerPort; intCurrentBalancerUplinkPort - 16 <= intCrossPortsOnEachBalancer; intCurrentBalancerUplinkPort++)
                            {
                                if (boolEshelon && intCurrentBalancerUplinkPort - 16 > intCrossPortsOnEachBalancer / 2) intBalancerRecalculatedPort = intCurrentBalancerUplinkPort + 8 - intCrossPortsOnEachBalancer / 2;
                                else intBalancerRecalculatedPort = intCurrentBalancerUplinkPort;
                                strCurrentPortName = "p" + intBalancerRecalculatedPort;
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort].Data3 = Convert.ToString(intCurrentBalancerFrame);
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort].Text = strCurrentPortName;
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                                if (arrShapesBypass100MonPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort] != null)
                                {
                                    //Draw Line And Fake Rectangle
                                    arrShapesBalancerFakeUplinkLines[intCurrentBalancerFrame, intBalancerRecalculatedPort] = page1.DrawLine(doubNextPortStartPointX - 1.5 - 0.2 * (intCurrentBalancerUplinkPort - 16) - intCurrentBalancerFrame * 4, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4);
                                    arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort] = page1.DrawOval(doubNextPortStartPointX - 1.1 - 0.4 * (intCurrentBalancerUplinkPort % 2), doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.7 - 0.4 * (intCurrentBalancerUplinkPort % 2), doubNextPortStartPointY - 0.6);
                                    arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                    arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort].Text = arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, 0]["Cable_Number"];
                                    arrShapesBalancerFakeUplinkConnections[intCurrentBalancerFrame, intBalancerRecalculatedPort] = page1.DrawRectangle(doubNextPortStartPointX - 1.5 - 0.2 * (intCurrentBalancerUplinkPort - 16) - intCurrentBalancerFrame * 4, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 1.5 - 0.2 * (intCurrentBalancerUplinkPort - 16) - intCurrentBalancerFrame * 4, doubNextPortStartPointY - 0.4);
                                    arrShapesBalancerFakeUplinkConnections[intCurrentBalancerFrame, intBalancerRecalculatedPort].AutoConnect(arrShapesBypass100MonFakeConnection[intCurrentBalancerFrame, intBalancerRecalculatedPort], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                                    //////////  MySQL
                                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 15, 1672);";
                                    command.ExecuteNonQuery();
                                    strNeighborPortId = arrShapesBypass100MonPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort].Data1;
                                    strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                                    command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                                    command.ExecuteNonQuery();

                                    //КЖ
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, 0].Add("Device_B_Name", strCurrentDeviceHostname);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, 0].Add("Port_B_Name", strCurrentPortName);
                                    list_CableJournal_Bypass_Balancer.Add(new Dictionary<string, string>());
                                    foreach (string key in arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, 0].Keys)
                                    {
                                        list_CableJournal_Bypass_Balancer[list_CableJournal_Bypass_Balancer.Count - 1].Add(key, arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, 0][key]);
                                    };
                                };
                                if (intCurrentBalancerUplinkPort % 2 == 0 && !boolEshelon) doubNextPortStartPointY -= 0.4;
                                else doubNextPortStartPointY -= 0.2;

                                //if (boolEshelon && intCurrentBalancerUplinkPort - 16 == intCrossPortsOnEachBalancer / 2) doubNextPortStartPointY -= 0.6;
                                if (boolEshelon && intCurrentBalancerUplinkPort - 16 == intCrossPortsOnEachBalancer / 2) doubNextPortStartPointY -= 4;              //подъём

                            };
                            intStartBalancerPort = 17;
                            intCurrentHighwayPeremyckaNumber = intCrossPortsOnEachBalancer / 2;
                        };

                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw ELB to 10G Uplink Ports (IS40 & IBS1UP)   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        Console.WriteLine($"Портов на балансере {intCurrentBalancerFrame}: {arrUplinkPortsOnBalancer[intCurrentBalancerFrame]}");
                        Console.WriteLine($"Гидр, посчитанных на IS40: {intBypassIs40HydrasTotal}");
                        if (intBypassIs40HydrasTotal > 0)
                        {
                            Console.WriteLine($"Балансер: {intCurrentBalancerFrame}, Последний порт: {arrUplinkPortsOnBalancer[intCurrentBalancerFrame]}.");
                            if (intBypassIs40HydrasTotal - intUsedUplinkHydrasCounter > intMaximumUplinkPortsOnBalancer) intLastUplinkPortOnBalancer = intMaximBalancerPortNumber;
                            else intLastUplinkPortOnBalancer = 16 + intBypassIs40HydrasTotal - intUsedUplinkHydrasCounter;
                            for (int intCurrentBalancerUplinkPort = intStartBalancerPort; intCurrentBalancerUplinkPort <= arrUplinkPortsOnBalancer[intCurrentBalancerFrame]; intCurrentBalancerUplinkPort++)
                            {
                                if (boolEshelon && intCurrentBalancerUplinkPort - 16 > (intLastUplinkPortOnBalancer - 16) / 2) intBalancerRecalculatedPort = intCurrentBalancerUplinkPort + 8 - (intLastUplinkPortOnBalancer - 16) / 2;
                                else intBalancerRecalculatedPort = intCurrentBalancerUplinkPort;
                                intUsedUplinkHydrasCounter++;
                                strCurrentPortName = "p" + intBalancerRecalculatedPort;
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort].Data3 = Convert.ToString(intCurrentBalancerFrame);
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort].Text = strCurrentPortName;
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                                //Console.WriteLine($"Балансер: {intCurrentBalancerFrame}, Порт: {intBalancerRecalculatedPort}.");
                                if (arrShapesBypass100MonFakeConnection[intCurrentBalancerFrame, intBalancerRecalculatedPort] != null)
                                {
                                    //Draw Line And Fake Rectangle
                                    arrShapesBalancerFakeUplinkLines[intCurrentBalancerFrame, intBalancerRecalculatedPort] = page1.DrawLine(doubNextPortStartPointX - 1.5 - 0.2 * (intCurrentBalancerUplinkPort - 16) - intCurrentBalancerFrame * 4, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4);
                                    arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort] = page1.DrawOval(doubNextPortStartPointX - 1.1 - 0.2 * (intCurrentBalancerUplinkPort % 2), doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.7 - 0.2 * (intCurrentBalancerUplinkPort % 2), doubNextPortStartPointY - 0.6);
                                    arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                    arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort].Text = arrShapesBypassMonFakeCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort].Text;
                                    arrShapesBalancerFakeUplinkConnections[intCurrentBalancerFrame, intBalancerRecalculatedPort] = page1.DrawRectangle(doubNextPortStartPointX - 1.5 - 0.2 * (intCurrentBalancerUplinkPort - 16) - intCurrentBalancerFrame * 4, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 1.5 - 0.2 * (intCurrentBalancerUplinkPort - 16) - intCurrentBalancerFrame * 4, doubNextPortStartPointY - 0.4);
                                    arrShapesBalancerFakeUplinkConnections[intCurrentBalancerFrame, intBalancerRecalculatedPort].AutoConnect(arrShapesBypass100MonFakeConnection[intCurrentBalancerFrame, intBalancerRecalculatedPort], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                                    //КЖ
                                    for (int inCurrentHydraEnd = 1; inCurrentHydraEnd <= 4; inCurrentHydraEnd++)
                                    {
                                        if (arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd] != null)
                                        {
                                            Console.Write($"Балансер: {intCurrentBalancerFrame}, Порт: {intBalancerRecalculatedPort}, Гидра: {inCurrentHydraEnd}");
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Device_B_Name", strCurrentDeviceHostname);
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Port_B_Name", strCurrentPortName + "-" + inCurrentHydraEnd);
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Cable_Number", arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort].Text);
                                            //Console.WriteLine($"   Записано.");
                                            //////////  MySQL
                                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName + "-" + inCurrentHydraEnd}', 9, 36);";
                                            command.ExecuteNonQuery();
                                            strNeighborPortId = arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd]["Cable_ID"];
                                            strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                                            command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                                            command.ExecuteNonQuery();

                                        }
                                        else
                                        {
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd] = new Dictionary<string, string>();
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Row", "");
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Cable_Name", "");
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Port_A_Name", "пусто");
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Device_A_Name", "");
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Cable_Type", "");
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Device_B_Name", "");
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Port_B_Name", "пусто");
                                            arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Add("Cable_Number", "");
                                        };


                                        list_CableJournal_Bypass_Balancer.Add(new Dictionary<string, string>());

                                        foreach (string key in arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd].Keys)
                                        {
                                            list_CableJournal_Bypass_Balancer[list_CableJournal_Bypass_Balancer.Count - 1].Add(key, arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort, inCurrentHydraEnd][key]);
                                        };
                                    };

                                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                                    if (boolEshelon)
                                    {
                                        strCurrentPortName = "p" + Convert.ToString(intBalancerRecalculatedPort + 8);
                                        doubShiftY = doubNextPortStartPointY - 5;      // здеся
                                        //  LAN-порты ---------------------------   Не отлажено!
                                        arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubShiftY - 0.5, doubNextPortStartPointX, doubShiftY - 0.3);
                                        arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8].Data3 = Convert.ToString(intCurrentBalancerFrame);
                                        arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8].Text = strCurrentPortName;
                                        arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";
                                        //Draw Line And Fake Rectangle
                                        arrShapesBalancerFakeUplinkLines[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8] = page1.DrawLine(doubNextPortStartPointX - 1.5 - 0.2 * (intCurrentBalancerUplinkPort - 16) - intCurrentBalancerFrame * 4, doubShiftY - 0.4, doubNextPortStartPointX - 0.5, doubShiftY - 0.4);
                                        arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8] = page1.DrawOval(doubNextPortStartPointX - 1.1 - 0.2 * (intCurrentBalancerUplinkPort % 2), doubShiftY - 0.2, doubNextPortStartPointX - 0.7 - 0.2 * (intCurrentBalancerUplinkPort % 2), doubShiftY - 0.6);
                                        arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                        arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8].Text = arrShapesBypassMonFakeCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8].Text;
                                        arrShapesBalancerFakeUplinkConnections[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8] = page1.DrawRectangle(doubNextPortStartPointX - 1.5 - 0.2 * (intCurrentBalancerUplinkPort - 16) - intCurrentBalancerFrame * 4, doubShiftY - 0.4, doubNextPortStartPointX - 1.5 - 0.2 * (intCurrentBalancerUplinkPort - 16) - intCurrentBalancerFrame * 4, doubShiftY - 0.4);



                                        Console.WriteLine($"Балансировщик {intCurrentBalancerFrame}, порт: {intBalancerRecalculatedPort + 8}");

                                        if (arrShapesBypass100MonFakeConnection[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8] != null)
                                        {
                                            arrShapesBalancerFakeUplinkConnections[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8].AutoConnect(arrShapesBypass100MonFakeConnection[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                            Console.WriteLine("Зацепили!");
                                        }
                                           



                                        //КЖ
                                        for (int inCurrentHydraEnd = 1; inCurrentHydraEnd <= 4; inCurrentHydraEnd++)
                                        {
                                            if (arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd] != null)
                                            {
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Device_B_Name", strCurrentDeviceHostname);
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Port_B_Name", strCurrentPortName + "-" + inCurrentHydraEnd);
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Cable_Number", arrShapesBalancerFakeUplinkCircles[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8].Text);

                                                //////////  MySQL
                                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName + "-" + inCurrentHydraEnd}', 9, 36);";
                                                command.ExecuteNonQuery();
                                                strNeighborPortId = arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd]["Cable_ID"];
                                                strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                                                command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                                                command.ExecuteNonQuery();

                                            }
                                            else
                                            {
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd] = new Dictionary<string, string>();
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Row", "");
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Cable_Name", "");
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Port_A_Name", "пусто");
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Device_A_Name", "");
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Cable_Type", "");
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Device_B_Name", "");
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Port_B_Name", "пусто");
                                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Add("Cable_Number", "");
                                            };


                                            list_CableJournal_Bypass_Balancer.Add(new Dictionary<string, string>());

                                            foreach (string key in arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd].Keys)
                                            {
                                                list_CableJournal_Bypass_Balancer[list_CableJournal_Bypass_Balancer.Count - 1].Add(key, arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intBalancerRecalculatedPort + 8, inCurrentHydraEnd][key]);
                                            };
                                        };

                                    };


                                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                }
                                else
                                {
                                    //Console.WriteLine($"Балансер: {intCurrentBalancerFrame}, Порт: {intBalancerRecalculatedPort}. Не найдена пара!");
                                };

                                doubNextPortStartPointY -= 0.4;
                            };
                            //if (arrUplinkPortsOnBalancer[intCurrentBalancerFrame] <= 18 || noMoreTwoLinksInHydra) intCurrentHighwayPeremyckaNumber = 1;
                            if (arrUplinkPortsOnBalancer[intCurrentBalancerFrame] <= 18) intCurrentHighwayPeremyckaNumber = 1;
                            else intCurrentHighwayPeremyckaNumber = 2;
                            //intCurrentHighwayPeremyckaNumber = CalculateDevicesQuantity(arrUplinkPortsOnBalancer[intCurrentBalancerFrame] - 16, 4);
                        };

                        if (boolEshelon)
                        {
                            //Console.WriteLine($"Перемычек на хайвэй {intCurrentBalancerFrame}: {intCurrentHighwayPeremyckaNumber}");
                            doubNextPortStartPointY -= 0.2;
                            //doubShiftY = doubNextPortStartPointY;
                            doubShiftY = doubNextPortStartPointY - 1;               //подъём; для десяток и соток разные значения подъёма нужны; как это сделать?
                            intCurrentHighwayCurrentPeremychkaPort = 33;
                            //if (intCurrentBalancerFrame == 1) intCurrentHighwayPeremyckaNumber = 4; else intCurrentHighwayPeremyckaNumber = 3;                    //Правим тут!
                            for (int inCurrentPeremychka = 1; inCurrentPeremychka <= intCurrentHighwayPeremyckaNumber; inCurrentPeremychka++)
                            {
                                intOverallPeremychkaNumber++;

                                //Перемычка. Верхний порт.
                                intCurrentHighwayCurrentPeremychkaPort--;
                                strCurrentPortName = "p" + Convert.ToString(intCurrentHighwayCurrentPeremychkaPort - 8);
                                arrShapesBalancerPeremychkaPorts[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubShiftY - 0.5, doubNextPortStartPointX, doubShiftY - 0.3);
                                arrShapesBalancerPeremychkaPorts[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8].Text = strCurrentPortName;
                                arrShapesBalancerPeremychkaPorts[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";
                                arrShapesBalancerPeremychkaLines[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8] = page1.DrawLine(doubNextPortStartPointX - 1.5 - 0.2 * inCurrentPeremychka - intCurrentBalancerFrame * 2, doubShiftY - 0.4, doubNextPortStartPointX - 0.5, doubShiftY - 0.4);
                                arrShapesBalancerPeremychkaLines[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8].CellsU["LineColor"].FormulaForceU = "THEMEGUARD(RGB(139,0,139))";
                                arrShapesBalancerFakePeremychkaCircles[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8] = page1.DrawOval(doubNextPortStartPointX - 1.1 - 0.2 * (intCurrentHighwayCurrentPeremychkaPort % 2), doubShiftY - 0.2, doubNextPortStartPointX - 0.7 - 0.2 * (intCurrentHighwayCurrentPeremychkaPort % 2), doubShiftY - 0.6);
                                arrShapesBalancerFakePeremychkaCircles[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesBalancerFakePeremychkaCircles[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8].Text = $"{intOverallPeremychkaNumber} Пер";
                                arrShapesBalancerFakePeremychkaConnections[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8] = page1.DrawRectangle(doubNextPortStartPointX - 1.5 - 0.2 * inCurrentPeremychka - intCurrentBalancerFrame * 2, doubShiftY - 0.4, doubNextPortStartPointX - 1.5 - 0.2 * inCurrentPeremychka - intCurrentBalancerFrame * 2, doubShiftY - 0.4);
                                arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort] = new Dictionary<string, string>();
                                arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Add("Device_A_Name", strCurrentDeviceHostname);
                                arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Add("Port_A_Name", strCurrentPortName);
                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 15, 1672);";
                                command.ExecuteNonQuery();
                                strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));

                                //Перемычка. Нижний порт.
                                strCurrentPortName = "p" + Convert.ToString(intCurrentHighwayCurrentPeremychkaPort);
                                arrShapesBalancerPeremychkaPorts[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubShiftY - 5.5, doubNextPortStartPointX, doubShiftY - 5.3);
                                arrShapesBalancerPeremychkaPorts[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Text = strCurrentPortName;
                                arrShapesBalancerPeremychkaPorts[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";
                                arrShapesBalancerPeremychkaLines[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort] = page1.DrawLine(doubNextPortStartPointX - 1.5 - 0.2 * inCurrentPeremychka - intCurrentBalancerFrame * 2, doubShiftY - 5.4, doubNextPortStartPointX - 0.5, doubShiftY - 5.4);
                                arrShapesBalancerPeremychkaLines[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].CellsU["LineColor"].FormulaForceU = "THEMEGUARD(RGB(139,0,139))";
                                arrShapesBalancerFakePeremychkaCircles[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort] = page1.DrawOval(doubNextPortStartPointX - 1.1 - 0.2 * (intCurrentHighwayCurrentPeremychkaPort % 2), doubShiftY - 5.2, doubNextPortStartPointX - 0.7 - 0.2 * (intCurrentHighwayCurrentPeremychkaPort % 2), doubShiftY - 5.6);
                                arrShapesBalancerFakePeremychkaCircles[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesBalancerFakePeremychkaCircles[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Text = $"{intOverallPeremychkaNumber} Пер";
                                arrShapesBalancerFakePeremychkaConnections[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort] = page1.DrawRectangle(doubNextPortStartPointX - 1.5 - 0.2 * inCurrentPeremychka - intCurrentBalancerFrame * 2, doubShiftY - 5.4, doubNextPortStartPointX - 1.5 - 0.2 * inCurrentPeremychka - intCurrentBalancerFrame * 2, doubShiftY - 5.4);
                                arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Add("Device_B_Name", strCurrentDeviceHostname);
                                arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Add("Port_B_Name", strCurrentPortName);
                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 15, 1672);";
                                command.ExecuteNonQuery();
                                strNeighborPortId = Convert.ToString(LastUsedId(command, "Port"));

                                //Коммутация верхнего и нижнего портов перемычки.
                                arrShapesBalancerFakePeremychkaConnections[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].AutoConnect(arrShapesBalancerFakePeremychkaConnections[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort - 8], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                                command.ExecuteNonQuery();

                                //КЖ
                                arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Add("Cable_Number", Convert.ToString(intOverallPeremychkaNumber) + " Пер");
                                arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Add("Cable_Name", $"{strCurrentDeviceHostname} --- {strCurrentDeviceHostname}");
                                arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Add("Cable_Type", "NR-QSFP-100G-AOC 100G QSFP28 <=> QSFP28");
                                list_CableJournal_Highway_Peremychka.Add(new Dictionary<string, string>());
                                foreach (string key in arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort].Keys)
                                {
                                    list_CableJournal_Highway_Peremychka[list_CableJournal_Highway_Peremychka.Count - 1].Add(key, arrCableJournal_Highway_Peremychka[intCurrentBalancerFrame, intCurrentHighwayCurrentPeremychkaPort][key]);
                                };

                                //////////  MySQL


                                doubShiftY -= 0.4;
                            };
                        };
                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw Downlink Ports (для прямой и крестовой)   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        intFilterPointerCross = 0;
                        intHydraStartPoint = intHydrasQuantityBwOnePair * (intCurrentBalancerFrame - 1);
                        intHydraPointerCross = intHydraStartPoint;

                        //strAssimetry = "да";                                                                    // Удалить!!!!!!!!!!!!!
                        //~~~~~~~~~~~~~~~~~~~~~~~~  Крестовая   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        if (boolCrossLayout || strAssimetry == "да" || strAssimetry == "эшкрест")
                        {
                            intSubseqHydrasCounter = 0;
                            for (int intCurrentBalancerDownlinkPort = 1; intCurrentBalancerDownlinkPort <= intTotalFilterHydrasQuantity / intTotalBalancersQuantity; intCurrentBalancerDownlinkPort++)
                            {
                                //Логика сдвига указателя фильтров и гидр одинакова.
                                //if (strAssimetry == "да") Console.WriteLine("Асимметрия. Крестовые соединения к балансерам.");
                                strCurrentPortName = "p" + intCurrentBalancerDownlinkPort;
                                if (intCurrentBalancerDownlinkPort > 16) strCurrentPortName = "p" + (49 - intCurrentBalancerDownlinkPort);
                                if (intHydraPointerCross == intHydraStartPoint) intFilterPointerCross++;
                                intHydraPointerCross++;
                                intSubseqHydrasCounter++;
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerCross, intHydraPointerCross] = page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 0.5, doubNextPortStartPointY - 0.3);
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerCross, intHydraPointerCross].Data3 = Convert.ToString(intCurrentBalancerFrame);
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerCross, intHydraPointerCross].Text = strCurrentPortName;
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerCross, intHydraPointerCross].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                                for (int intCurrentPortInHydra = 1; intCurrentPortInHydra <= 4; intCurrentPortInHydra++)
                                {
                                    //КЖ "Балансировщики - Фильтры"
                                    if (intCurrentPortInHydra == 1) intGlobalCableCounter++;
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra] = new Dictionary<string, string>();
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + " БФ");
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Cable_Name", $"{strCurrentDeviceHostname} --- {strFilterModel} ({intFilterPointerCross})");
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Port_A_Name", strCurrentPortName + "-" + intCurrentPortInHydra);
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Device_A_Name", strCurrentDeviceHostname);

                                    //////////  MySQL
                                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}-{intCurrentPortInHydra}', 9, 36);";
                                    command.ExecuteNonQuery();
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));
                                    //arrShapesBalancer1040DownlinkPorts[intFilterPointerCross, intHydraPointerCross].Data1 = Convert.ToString(LastUsedId(command, "Port"));
                                };


                                //Draw Line And Fake Rectangle
                                arrShapesBalancerFakeDownlinkLines[intFilterPointerCross, intHydraPointerCross] = page1.DrawLine(doubNextPortStartPointX + 0.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 1.5 + 0.2 * intCurrentBalancerDownlinkPort, doubNextPortStartPointY - 0.4);
                                arrShapesBalancerFakeDownlinkCircles[intFilterPointerCross, intHydraPointerCross] = page1.DrawOval(doubNextPortStartPointX + 1 + 0.2 * (intCurrentBalancerDownlinkPort % 2), doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 1.4 + 0.2 * (intCurrentBalancerDownlinkPort % 2), doubNextPortStartPointY - 0.6);
                                arrShapesBalancerFakeDownlinkCircles[intFilterPointerCross, intHydraPointerCross].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesBalancerFakeDownlinkCircles[intFilterPointerCross, intHydraPointerCross].Text = Convert.ToString(intGlobalCableCounter) + " БФ";
                                arrShapesBalancerFakeDownlinkConnections[intFilterPointerCross, intHydraPointerCross] = page1.DrawRectangle(doubNextPortStartPointX + 1.5 + 0.2 * intCurrentBalancerDownlinkPort, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 1.5 + 0.2 * intCurrentBalancerDownlinkPort, doubNextPortStartPointY - 0.4);

                                //Добавляем в спецификацию
                                list_Specification.Add("NR-QSFP-4X10G-AOC 40G QSFP");
                                doubNextPortStartPointY -= 0.4;



                                // Если текущий номер гидры совпал с количеством стыков балансера на один фильтр, сдвигаем указатель в начало .
                                if (intSubseqHydrasCounter == intHydrasQuantityBwOnePair)
                                {
                                    intHydraPointerCross = intHydraStartPoint;
                                    intSubseqHydrasCounter = 0;
                                };
                            };
                        }
                        //~~~~~~~~~~~~~~~~~~~~~~~~  Прямая   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        else
                        {
                            if (intTotalFilterHydrasQuantity < 16) intLastDownlinkPortOnBalancer = intTotalFilterHydrasQuantity;     //Если балансер единственный, все гидры к фильтрам приходят к нему
                            //else if (intTotalFilterHydrasQuantity - intUsedDownlinkHydrasCounter > intHydrasOnFilter) intLastDownlinkPortOnBalancer = 10; //изменить
                            else intLastDownlinkPortOnBalancer = intTotalFilterHydrasQuantity - intUsedDownlinkHydrasCounter;

                            //intLastDownlinkPortOnBalancer = 12;                     //Удалить!!!!!!!! Только для 5880.
                            //Console.WriteLine($"Балансер {intCurrentBalancerFrame}, Последний порт: {intLastDownlinkPortOnBalancer}.");
                            if (intTotalBalancers == 1) arrUplinkPortsOnBalancer[1] = intLastDownlinkPortOnBalancer + 16;

                            for (int intCurrentBalancerDownlinkPort = 1; intCurrentBalancerDownlinkPort <= intLastDownlinkPortOnBalancer; intCurrentBalancerDownlinkPort++)        //???????????????????????????????????? Было 12!
                            {
                                strCurrentPortName = "p" + intCurrentBalancerDownlinkPort;
                                if (intHydraPointerStraight == 0) intFilterPointerStraight++;
                                intHydraPointerStraight++;
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerStraight, intHydraPointerStraight] = page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 0.5, doubNextPortStartPointY - 0.3);
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerStraight, intHydraPointerStraight].Data3 = Convert.ToString(intCurrentBalancerFrame);
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerStraight, intHydraPointerStraight].Text = strCurrentPortName;
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerStraight, intHydraPointerStraight].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";
                                for (int intCurrentPortInHydra = 1; intCurrentPortInHydra <= 4; intCurrentPortInHydra++)
                                {
                                    if (intCurrentPortInHydra == 1) intGlobalCableCounter++;
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra] = new Dictionary<string, string>();
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + " БФ");
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Cable_Name", $"{strCurrentDeviceHostname} --- {strFilterModel} ({intFilterPointerStraight})");
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Port_A_Name", strCurrentPortName + "-" + intCurrentPortInHydra);
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Device_A_Name", strCurrentDeviceHostname);

                                    //////////  MySQL
                                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}-{intCurrentPortInHydra}', 9, 36);";
                                    command.ExecuteNonQuery();
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Cable_ID", Convert.ToString(LastUsedId(command, "Port")));

                                };

                                arrShapesBalancerFakeDownlinkLines[intFilterPointerStraight, intHydraPointerStraight] = page1.DrawLine(doubNextPortStartPointX + 0.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 1.5 + 0.2 * intCurrentBalancerDownlinkPort, doubNextPortStartPointY - 0.4);
                                arrShapesBalancerFakeDownlinkCircles[intFilterPointerCross, intHydraPointerCross] = page1.DrawOval(doubNextPortStartPointX + 1 + 0.2 * (intCurrentBalancerDownlinkPort % 2), doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 1.4 + 0.2 * (intCurrentBalancerDownlinkPort % 2), doubNextPortStartPointY - 0.6);
                                arrShapesBalancerFakeDownlinkCircles[intFilterPointerCross, intHydraPointerCross].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesBalancerFakeDownlinkCircles[intFilterPointerCross, intHydraPointerCross].Text = Convert.ToString(intGlobalCableCounter) + " БФ";
                                arrShapesBalancerFakeDownlinkConnections[intFilterPointerStraight, intHydraPointerStraight] = page1.DrawRectangle(doubNextPortStartPointX + 1.5 + 0.2 * intCurrentBalancerDownlinkPort, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 1.5 + 0.2 * intCurrentBalancerDownlinkPort, doubNextPortStartPointY - 0.4);

                                //Добавляем в спецификацию
                                list_Specification.Add("NR-QSFP-4X10G-AOC 40G QSFP");


                                if (intHydraPointerStraight == intHydrasOnFilter) intHydraPointerStraight = 0;

                                doubNextPortStartPointY -= 0.4;



                            };



                        };


                        doubStartPointNextShapeY -= intGapBwBalancers;

                        doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                        doubNextPortStartPointY = doubStartPointNextShapeY;
                        doubNextPortStartPointY = doubStartPointNextShapeY + intGapBwBalancers;

                    };
                };

                int intFiltersNumberFromPorts = listShapesBalancerDownlinkPorts.Count / 4;
                int intTotalFiltersFinal = Math.Max(intTotalFiltersFromBw, intFiltersNumberFromPorts);
                intTotalFiltersFinal = intTotalFiltersFromBw;
                int intMaximumHydras = intTotalFiltersFinal * 4;



                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////////////////////////////////////  Draw Filters ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                intCurrentRackSlot = 10;


                List<Visio.Shape> listDeviceLogPorts = new List<Visio.Shape>();                            //Массив для LOG-портов фильтров
                List<Visio.Shape> listDeviceLogCircles = new List<Visio.Shape>();
                List<Visio.Shape> listDeviceLogLines = new List<Visio.Shape>();
                List<Visio.Shape> listDeviceLogFakeRects = new List<Visio.Shape>();

                Visio.Shape[,] arrShapesFilterFakeUplinkLines = new Visio.Shape[50, 100];
                Visio.Shape[,] arrShapesFilterFakeUplinkConnections = new Visio.Shape[50, 100];
                Visio.Shape[,] arrShapesFilterFakeUplinkCircles = new Visio.Shape[50, 100];


                doubStartPointNextShapeX += 15 + 0.8 * intTotalFiltersQuantity;
                doubStartPointNextShapeY = doubUpperStartPoint;

                List<Visio.Shape> listShapesFilter4160Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesFilterPorts = new List<Visio.Shape>();
                List<Visio.Shape> listFilterHydraConnectors = new List<Visio.Shape>();
                List<Visio.Shape> listShapesSingleFilterPorts = new List<Visio.Shape>();


                int intCurrentHydraOnFilter;

                if (!boolNoBalancer)
                {
                    for (int intCurrentFilterFrame = 1; intCurrentFilterFrame <= intTotalFiltersQuantity; intCurrentFilterFrame++)
                    {
                        //--------------------------------------------------  New Filter Chassis Start ---------------------------------------------

                        arrShapesFilterDevices[intCurrentFilterFrame] = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 6);
                        strCurrentDeviceHostname = "Filter-" + strFilterModel + " (" + intCurrentFilterFrame + ")";
                        strNameForRackTables = $"{strFullSiteIndex}-F{strFilterModel}-{intCurrentFilterFrame}";
                        arrShapesFilterDevices[intCurrentFilterFrame].Text = strCurrentDeviceHostname;
                        arrShapesFilterDevices[intCurrentFilterFrame].Data3 = strCurrentDeviceHostname;
                        arrShapesFilterDevices[intCurrentFilterFrame].get_Cells("FillForegnd").FormulaU = "=RGB(175,238,238)";

                        intCurrentHydraOnFilter = 0;

                        /////////////// Добавляем шасси в БД
                        intCurrentRackSlot++;
                        intDeviceId = FillRackSlot(command, strObjectIndex, intFilterRackId, intCurrentRackSlot, 50002, strNameForRackTables);


                        //--------------------------------------------------  New Filter Chassis End ---------------------------------------------

                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MNG";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MNG";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = strCurrentDeviceHostname;
                        // MySQL Mgmt.
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'MNG', 1, 24);";
                        command.ExecuteNonQuery();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));
                        //Mgmt Lines
                        listDeviceMgmtLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.7));
                        listDeviceMgmtCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.6, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1));
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                        listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.7, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.7));
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "MNG";
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data3 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 0.3));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "IPMI";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "IPMI";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = strCurrentDeviceHostname;
                        // MySQL Mgmt.
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'IPMI', 1, 24);";
                        command.ExecuteNonQuery();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));
                        //Mgmt Lines
                        listDeviceMgmtLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.6));
                        listDeviceMgmtCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 1.4));
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                        listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.6, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.6));
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "IPMI";
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data3 = strCurrentDeviceHostname;

                        //Порт SP1
                        listDeviceLogPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 1.1, doubStartPointNextShapeY + 0.3));
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Text = "SP1/LOG2";
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Rotate90();
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data2 = "SP1/LOG2";
                        //listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data3 = strCurrentDeviceHostname;
                        //Log Lines
                        listDeviceLogLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 1.5));
                        listDeviceLogCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.6, doubStartPointNextShapeX + 1.1, doubStartPointNextShapeY + 1));
                        listDeviceLogCircles[listDeviceLogCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        listDeviceLogCircles[listDeviceLogCircles.Count - 1].Text = listDeviceLogCircles.Count + " LG";
                        listDeviceLogFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 1.5, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 1.5));
                        // MySQL
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'SP1', 9, 36);";
                        command.ExecuteNonQuery();
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));
                        //Console.WriteLine($"PortCount: {listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data3}.");

                        //---------------------  New Filter Ports Start ----------------------------------------

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        listHydraLines.Clear();
                        // Логика определения портов для прямой и крестовой схем производится на стороне балансировщиков.
                        for (int intCurrentFilterPort = 1; intCurrentFilterPort <= 4 * intHydrasOnFilter; intCurrentFilterPort++)
                        {
                            strCurrentPortName = "Te" + intCurrentFilterPort;
                            arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                            arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort].Data3 = Convert.ToString(intCurrentFilterFrame);
                            arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort].Text = strCurrentPortName;
                            arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";


                            listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX - 0.7, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4));
                            if (listHydraLines.Count == 1) intCurrentHydraOnFilter++;

                            //////////////////////////////////////////////
                            arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, listHydraLines.Count].Add("Port_B_Name", strCurrentPortName + " (AOC c" + listHydraLines.Count + ")");
                            arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, listHydraLines.Count].Add("Device_B_Name", strCurrentDeviceHostname);
                            list_CableJournal_Balancer_Filter.Add(new Dictionary<string, string>());
                            foreach (string key in arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, listHydraLines.Count].Keys)
                            {
                                list_CableJournal_Balancer_Filter[list_CableJournal_Balancer_Filter.Count - 1].Add(key, arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, listHydraLines.Count][key]);
                            };

                            //////////  MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                            command.ExecuteNonQuery();
                            strNeighborPortId = arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, listHydraLines.Count]["Cable_ID"];//arrShapesBalancer1040DownlinkPorts[intCurrentFilterFrame, intCurrentHydraOnFilter].Data1;
                            strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                            command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                            command.ExecuteNonQuery();


                            //После каждого 4го нарисованного порта объединяем конструкцию в гидру. Гидра становится объектом, к ним цепляются порты балансировщиков.
                            if (listHydraLines.Count == 4)
                            {
                                listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX - 0.7, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.7, doubNextPortStartPointY + 0.5));
                                vsoWindow.DeselectAll();
                                foreach (Visio.Shape objHydraSingleLine in listHydraLines)
                                {
                                    vsoWindow.Select(objHydraSingleLine, 2);
                                };
                                vsoSelection = vsoWindow.Selection;
                                arrFilterHydraConnectors[intCurrentFilterFrame, intCurrentHydraOnFilter] = vsoSelection.Group();
                                arrFilterHydraConnectors[intCurrentFilterFrame, intCurrentHydraOnFilter].Data1 = Convert.ToString(intCurrentFilterFrame);
                                listHydraLines.Clear();
                                arrShapesFilterFakeUplinkLines[intCurrentFilterFrame, intCurrentHydraOnFilter] = page1.DrawLine(doubNextPortStartPointX - 1.7 - (intTotalFiltersQuantity - intCurrentFilterFrame) - (intHydrasOnFilter - intCurrentHydraOnFilter) * 0.2, doubNextPortStartPointY, doubNextPortStartPointX - 0.7, doubNextPortStartPointY);
                                arrShapesFilterFakeUplinkCircles[intCurrentFilterFrame, intCurrentHydraOnFilter] = page1.DrawOval(doubNextPortStartPointX - 1.4, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 1, doubNextPortStartPointY + 0.2);
                                arrShapesFilterFakeUplinkCircles[intCurrentFilterFrame, intCurrentHydraOnFilter].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesFilterFakeUplinkCircles[intCurrentFilterFrame, intCurrentHydraOnFilter].Text = arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, 1]["Cable_Number"];
                                arrShapesFilterFakeUplinkConnections[intCurrentFilterFrame, intCurrentHydraOnFilter] = page1.DrawRectangle(doubNextPortStartPointX - 1.7 - (intTotalFiltersQuantity - intCurrentFilterFrame) - (intHydrasOnFilter - intCurrentHydraOnFilter) * 0.2, doubNextPortStartPointY, doubNextPortStartPointX - 1.7 - (intTotalFiltersQuantity - intCurrentFilterFrame) - (intHydrasOnFilter - intCurrentHydraOnFilter) * 0.2, doubNextPortStartPointY);
                                if (arrShapesBalancerFakeDownlinkConnections[intCurrentFilterFrame, intCurrentHydraOnFilter] != null) arrShapesFilterFakeUplinkConnections[intCurrentFilterFrame, intCurrentHydraOnFilter].AutoConnect(arrShapesBalancerFakeDownlinkConnections[intCurrentFilterFrame, intCurrentHydraOnFilter], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            };

                            if (intCurrentFilterPort % 2 > 0) doubNextPortStartPointY -= 0.2;
                            else doubNextPortStartPointY -= 0.5;
                        }

                        //---------------------  New Filter Ports End ----------------------------------------

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        doubStartPointNextShapeY -= 10;

                    };
                }
                ////////////////////////////////    Фильтр напрямую с байпасов   //////////////////////////
                else
                {
                    //Console.WriteLine($" LAN-линков: {intLinkCounter10Old + intLinkCounter10New}");

                    doubStartPointNextShapeY -= 1;
                    doubStartPointNextShapeX = doubBypassEndX + 8;

                    for (int intCurrentFilterFrame = 1; intCurrentFilterFrame <= intTotalFiltersQuantity; intCurrentFilterFrame++)
                    {
                        arrShapesFilterDevices[intCurrentFilterFrame] = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 0.4 * intPortsNumberOnSingleFilter);
                        strCurrentDeviceHostname = "Filter-" + strFilterModel + " (" + intCurrentFilterFrame + ")";
                        strNameForRackTables = $"{strFullSiteIndex}-F{strFilterModel}-{intCurrentFilterFrame}";
                        arrShapesFilterDevices[intCurrentFilterFrame].Text = strCurrentDeviceHostname;
                        arrShapesFilterDevices[intCurrentFilterFrame].get_Cells("FillForegnd").FormulaU = "=RGB(175,238,238)";

                        /////////////// Добавляем шасси в БД
                        intCurrentRackSlot++;
                        intDeviceId = FillRackSlot(command, strObjectIndex, intFilterRackId, intCurrentRackSlot, 50002, strNameForRackTables);

                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MNG";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MNG";
                        //Mgmt Lines
                        listDeviceMgmtLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.7));
                        listDeviceMgmtCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.6, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1));
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                        listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.7, doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1.7));
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "MNG";
                        // MySQL Mgmt.
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'MNG', 1, 24);";
                        command.ExecuteNonQuery();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 0.3));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "IPMI";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "IPMI";
                        //Mgmt Lines
                        listDeviceMgmtLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.5));
                        listDeviceMgmtCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 1, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 1.4));
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                        listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.5, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.5));
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "IPMI";
                        // MySQL Mgmt.
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'IPMI', 1, 24);";
                        command.ExecuteNonQuery();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                        listDeviceLogPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 1.1, doubStartPointNextShapeY + 0.3));
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Text = "SP1";
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Rotate90();
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data2 = "SP1";
                        //Log Lines
                        listDeviceLogLines.Add(page1.DrawLine(doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 1.5));
                        listDeviceLogCircles.Add(page1.DrawOval(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.6, doubStartPointNextShapeX + 1.1, doubStartPointNextShapeY + 1));
                        listDeviceLogCircles[listDeviceLogCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        listDeviceLogCircles[listDeviceLogCircles.Count - 1].Text = listDeviceLogCircles.Count + " LG";
                        listDeviceLogFakeRects.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 1.5, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 1.5));

                        // MySQL
                        command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'SP1', 9, 36);";
                        command.ExecuteNonQuery();
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        for (int intCurrentFilterPort = 1; intCurrentFilterPort <= intPortsNumberOnSingleFilter; intCurrentFilterPort++)
                        {
                            if (arrShapesBypass100MonFakeConnection[intCurrentFilterFrame, intCurrentFilterPort] != null)
                            {
                                strCurrentPortName = "Te " + intCurrentFilterPort;
                                arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1);
                                arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort].Text = strCurrentPortName;
                                arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";
                                //Fakes
                                arrShapesFilterFakeUplinkLines[intCurrentFilterFrame, intCurrentFilterPort] = page1.DrawLine(doubNextPortStartPointX - 5 + intCurrentFilterPort * 0.2, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.2);
                                arrShapesFilterFakeUplinkCircles[intCurrentFilterFrame, intCurrentFilterPort] = page1.DrawOval(doubNextPortStartPointX - 1 - 0.4 * (intCurrentFilterPort % 2), doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.6 - 0.4 * (intCurrentFilterPort % 2), doubNextPortStartPointY);
                                arrShapesFilterFakeUplinkCircles[intCurrentFilterFrame, intCurrentFilterPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                                arrShapesFilterFakeUplinkCircles[intCurrentFilterFrame, intCurrentFilterPort].Text = arr_CableJournal_Bypass_Balancer[intCurrentFilterFrame, intCurrentFilterPort, 0]["Cable_Number"];
                                arrShapesFilterFakeUplinkConnections[intCurrentFilterFrame, intCurrentFilterPort] = page1.DrawRectangle(doubNextPortStartPointX - 5 + intCurrentFilterPort * 0.2, doubNextPortStartPointY - 0.2, doubNextPortStartPointX - 5 + intCurrentFilterPort * 0.2, doubNextPortStartPointY - 0.2);
                                //Connect
                                arrShapesFilterFakeUplinkConnections[intCurrentFilterFrame, intCurrentFilterPort].AutoConnect(arrShapesBypass100MonFakeConnection[intCurrentFilterFrame, intCurrentFilterPort], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                if (intCurrentFilterPort % 2 > 0) doubNextPortStartPointY -= 0.2;
                                else doubNextPortStartPointY -= 0.5;

                                //  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                                //КЖ к байпасам
                                arr_CableJournal_Bypass_Balancer[intCurrentFilterFrame, intCurrentFilterPort, 0].Add("Device_B_Name", strCurrentDeviceHostname);

                                arr_CableJournal_Bypass_Balancer[intCurrentFilterFrame, intCurrentFilterPort, 0].Add("Port_B_Name", strCurrentPortName);
                                list_CableJournal_Bypass_Balancer.Add(new Dictionary<string, string>());
                                foreach (string key in arr_CableJournal_Bypass_Balancer[intCurrentFilterFrame, intCurrentFilterPort, 0].Keys)
                                {
                                    list_CableJournal_Bypass_Balancer[list_CableJournal_Bypass_Balancer.Count - 1].Add(key, arr_CableJournal_Bypass_Balancer[intCurrentFilterFrame, intCurrentFilterPort, 0][key]);
                                };

                                //////////  MySQL
                                //command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', 9, 36);";
                                //Console.WriteLine($"Filter: {strLinkTypeId}/{strLinkSubTypeId}");
                                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strCurrentPortName}', {strLinkTypeId}, {strLinkSubTypeId});";
                                command.ExecuteNonQuery();
                                strNeighborPortId = arr_CableJournal_Bypass_Balancer[intCurrentFilterFrame, intCurrentFilterPort, 0]["Cable_ID"];
                                strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                                //Console.WriteLine($"porta: {strNeighborPortId}, portb: {strLocalPortId}");
                                command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                                command.ExecuteNonQuery();
                            };
                        };
                        doubStartPointNextShapeY -= 10;
                    };

                };



                intGlobalCableCounter = 0;
                int intLast3348Port = 0;

                //////////  Draw Continent Router   //////////

                intCurrentRackSlot = 10;

                doubStartPointNextShapeX = doubNextPortStartPointX - 2;
                doubStartPointNextShapeY = intdoubTopLineY + 3 + 0.1 * listDeviceMgmtPorts.Count;
                double doubContStartX = doubStartPointNextShapeX;
                double doubContStartY = doubStartPointNextShapeY + 0.8;

                string strContinentUplinkPort;
                string strContinentToMes3348Port;
                string strContinentToMes5332aPort;
                string strContinentHostname;

                if (boolContinentIpcR300)
                {
                    strContinentHostname = $"IPC-R{strContinentModel}";
                    strContinentUplinkPort = "ix2";
                    strContinentToMes5332aPort = "ix3";
                    strContinentToMes3348Port = "igb0";
                    //Добавляем в спецификацию
                    list_Specification.Add("IPCR300");
                }
                else
                {
                    strContinentHostname = "IPC-100";
                    strContinentUplinkPort = "0 или 2";
                    strContinentToMes5332aPort = "1";
                    strContinentToMes3348Port = "3";
                    //Добавляем в спецификацию
                    list_Specification.Add("IPC100");
                };

                Console.WriteLine("Добавляем континент.");

                //Рисуем шасси континента
                Visio.Shape objContinentSwitch = page1.DrawRectangle(doubStartPointNextShapeX - 1, doubStartPointNextShapeY + 0.5, doubStartPointNextShapeX, doubStartPointNextShapeY + 2.2);
                strCurrentDeviceHostname = strContinentHostname;
                strNameForRackTables = $"{strFullSiteIndex}-{strContinentHostname}";
                objContinentSwitch.Text = strCurrentDeviceHostname;
                objContinentSwitch.get_Cells("FillForegnd").FormulaU = "=RGB(176,196,222)";
                /////////////// Добавляем шасси в БД
                intCurrentRackSlot++;
                intDeviceId = FillRackSlot(command, strObjectIndex, intNetSrvRackId, intCurrentRackSlot, 50024, strNameForRackTables);

                List<Visio.Shape> listShapesContinentPorts = new List<Visio.Shape>();
                doubNextPortStartPointX = doubStartPointNextShapeX - 3;
                doubNextPortStartPointY = doubStartPointNextShapeY + 0.2;

                //Рисуем аплинк-порт континента
                Visio.Shape visContinentUplinkPort = page1.DrawRectangle(doubStartPointNextShapeX - 1.4, doubStartPointNextShapeY + 0.7, doubStartPointNextShapeX - 1, doubStartPointNextShapeY + 0.9);
                visContinentUplinkPort.Text = strContinentUplinkPort;
                visContinentUplinkPort.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                // MySQL
                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strContinentUplinkPort}', 4, 1204);";
                command.ExecuteNonQuery();
                visContinentUplinkPort.Data3 = Convert.ToString(LastUsedId(command, "Port"));

                //Рисуем порт континента к MES5332A
                Visio.Shape visContinentToMes5332aPort = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY + 1.7, doubStartPointNextShapeX + 0.4, doubStartPointNextShapeY + 1.9);
                visContinentToMes5332aPort.Text = strContinentToMes5332aPort;
                visContinentToMes5332aPort.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                //IC Lines
                Visio.Shape visContinentToMes5332aLine = page1.DrawLine(doubStartPointNextShapeX + 0.4, doubStartPointNextShapeY + 1.8, doubStartPointNextShapeX + 1.2, doubStartPointNextShapeY + 1.8);
                Visio.Shape visContinentToMes5332aCircle = page1.DrawOval(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 1.6, doubStartPointNextShapeX + 1.1, doubStartPointNextShapeY + 2);
                visContinentToMes5332aCircle.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                Visio.Shape visContinentToMes5332aFakeRects = page1.DrawRectangle(doubStartPointNextShapeX + 1.2, doubStartPointNextShapeY + 1.8, doubStartPointNextShapeX + 1.2, doubStartPointNextShapeY + 1.8);

                // MySQL
                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strContinentToMes5332aPort}', 4, 1202);";
                command.ExecuteNonQuery();
                visContinentToMes5332aPort.Data3 = Convert.ToString(LastUsedId(command, "Port"));

                //Рисуем порт континента к MES3348
                Visio.Shape visContinentToMes3348Port = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY + 0.7, doubStartPointNextShapeX + 0.4, doubStartPointNextShapeY + 0.9);
                visContinentToMes3348Port.Text = strContinentToMes3348Port;
                visContinentToMes3348Port.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                //IC Lines
                Visio.Shape visContinentToMes3348Line = page1.DrawLine(doubStartPointNextShapeX + 0.4, doubStartPointNextShapeY + 0.8, doubStartPointNextShapeX + 1.2, doubStartPointNextShapeY + 0.8);
                Visio.Shape visContinentToMes3348Circle = page1.DrawOval(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.6, doubStartPointNextShapeX + 1.1, doubStartPointNextShapeY + 1);
                visContinentToMes3348Circle.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                Visio.Shape visContinentToMes3348FakeRect = page1.DrawRectangle(doubStartPointNextShapeX + 1.2, doubStartPointNextShapeY + 0.8, doubStartPointNextShapeX + 1.2, doubStartPointNextShapeY + 0.8);
                // MySQL
                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strContinentToMes3348Port}', 1, 24);";
                command.ExecuteNonQuery();
                visContinentToMes3348Port.Data3 = Convert.ToString(LastUsedId(command, "Port"));

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                doubNextPortStartPointX = doubStartPointNextShapeX + 4;
                doubNextPortStartPointY = doubStartPointNextShapeY + 3.3;

                //////////  Draw MES Log Switch   //////////
                Visio.Shape shapeMesLogDevice = page1.DrawRectangle(doubStartPointNextShapeX + 4, doubStartPointNextShapeY + 3.6, doubStartPointNextShapeX + 5 + listDeviceLogPorts.Count * 0.4, doubStartPointNextShapeY + 4.6 + intTotalLogServers * 0.1);
                strCurrentDeviceHostname = "MES5332A (1)";
                strNameForRackTables = $"{strFullSiteIndex}-LOGSW";
                shapeMesLogDevice.Text = strCurrentDeviceHostname;
                shapeMesLogDevice.get_Cells("FillForegnd").FormulaU = "=RGB(176,196,222)";
                //Добавляем в спецификацию
                list_Specification.Add("MES5332A");

                /////////////// Добавляем шасси в БД
                intCurrentRackSlot += 5;
                intDeviceId = FillRackSlot(command, strObjectIndex, intNetSrvRackId, intCurrentRackSlot, 8, strNameForRackTables);

                List<Visio.Shape> listShapesMesLogDownLinkPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesLogDownLinkLines = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesLogDownLinkCircles = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesLogDownLinkFakeRects = new List<Visio.Shape>();

                List<Visio.Shape> listShapesMesLogUpLinkPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesLogUpLinkLines = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesLogUpLinkCircles = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesLogUpLinkFakeRects = new List<Visio.Shape>();

                int intMes32Port = 0;

                //////////////////////  Add Log-MES Ports  //////////////////////////
                //Рисуем порты MES5332A на стык с LOG-портами фильтров
                for (int intCurrentMesPort = 1; intCurrentMesPort <= listDeviceLogPorts.Count; intCurrentMesPort++)
                {
                    intMes32Port++;
                    listShapesMesLogDownLinkPorts.Add(page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.4, doubNextPortStartPointY + 0.2));
                    listShapesMesLogDownLinkPorts[listShapesMesLogDownLinkPorts.Count - 1].Text = "XGE-" + intMes32Port;
                    doubNextPortStartPointX += 0.2;
                    listShapesMesLogDownLinkPorts[listShapesMesLogDownLinkPorts.Count - 1].Rotate90();
                    listShapesMesLogDownLinkPorts[listShapesMesLogDownLinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    //listShapesMesLogDownLinkPorts[listShapesMesLogDownLinkPorts.Count - 1].Data1 = "MES5332A";
                    listShapesMesLogDownLinkPorts[listShapesMesLogDownLinkPorts.Count - 1].Data1 = $"{strFullSiteIndex}-LOGSW";
                    listShapesMesLogDownLinkPorts[listShapesMesLogDownLinkPorts.Count - 1].Data2 = "XGE-" + intMes32Port;

                    //Log Lines
                    listShapesMesLogDownLinkLines.Add(page1.DrawLine(doubNextPortStartPointX, doubNextPortStartPointY - 0.1, doubNextPortStartPointX, doubNextPortStartPointY - 1.4));

                    listShapesMesLogDownLinkCircles.Add(page1.DrawOval(doubNextPortStartPointX - 0.2, doubNextPortStartPointY - 0.6 - 0.4 * (intCurrentMesPort % 2), doubNextPortStartPointX + 0.2, doubNextPortStartPointY - 0.2 - 0.4 * (intCurrentMesPort % 2)));
                    listShapesMesLogDownLinkCircles[listShapesMesLogDownLinkCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                    listShapesMesLogDownLinkCircles[listShapesMesLogDownLinkCircles.Count - 1].Text = listShapesMesLogDownLinkCircles.Count + " LG";
                    listShapesMesLogDownLinkFakeRects.Add(page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY - 1.4, doubNextPortStartPointX, doubNextPortStartPointY - 1.4));

                    listDeviceLogFakeRects[listShapesMesLogDownLinkFakeRects.Count - 1].AutoConnect(listShapesMesLogDownLinkFakeRects[listShapesMesLogDownLinkFakeRects.Count - 1], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                    listCableJournal_Log.Add(new Dictionary<string, string>());
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Number", listShapesMesLogDownLinkPorts.Count + " LG");
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Name", $"{strCurrentDeviceHostname} --- {strFilterModel} ({listShapesMesLogDownLinkPorts.Count})");
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Device_A", strCurrentDeviceHostname);
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Port_A", "XGE-" + listShapesMesLogDownLinkPorts.Count);
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Device_B", $"{strFilterModel} ({listShapesMesLogDownLinkPorts.Count})");
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Port_B", "SP1");
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Type", "NR-SFP-10G-AOC 10G SFP+ <=> SFP+ (10G, SFP+, AOC, 2м)");


                    // MySQL
                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'XGE-{intMes32Port}', 9, 36);";
                    command.ExecuteNonQuery();
                    strNeighborPortId = listDeviceLogPorts[listShapesMesLogDownLinkFakeRects.Count - 1].Data3;
                    strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                    //Console.WriteLine($"Local Port: {strLocalPortId}, Remote Port: {strNeighborPortId}");
                    command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                    command.ExecuteNonQuery();


                };

                doubNextPortStartPointX = doubStartPointNextShapeX + 5 + listDeviceLogPorts.Count * 0.4;
                doubNextPortStartPointY = doubStartPointNextShapeY + 3.9;


                //Рисуем на MES5332A порт OOB - на стык с последним в цепи MES3348
                Visio.Shape visMes5332aTo3348Port = page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 0.4, doubNextPortStartPointY);
                visMes5332aTo3348Port.Text = "OOB";
                visMes5332aTo3348Port.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                //IC Lines
                Visio.Shape visMes5332aTo3348Line = page1.DrawLine(doubNextPortStartPointX + 0.4, doubNextPortStartPointY - 0.1, doubNextPortStartPointX + 1.6, doubNextPortStartPointY - 0.1);
                Visio.Shape visMes5332aTo3348Circle = page1.DrawOval(doubNextPortStartPointX + 0.9, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 1.3, doubNextPortStartPointY + 0.1);
                visMes5332aTo3348Circle.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                Visio.Shape visMes5332aTo3348FakeRect = page1.DrawRectangle(doubNextPortStartPointX + 1.6, doubNextPortStartPointY - 0.1, doubNextPortStartPointX + 1.6, doubNextPortStartPointY - 0.1);
                // MySQL
                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'OOB', 1, 24);";
                command.ExecuteNonQuery();
                visMes5332aTo3348Port.Data3 = Convert.ToString(LastUsedId(command, "Port"));

                //Рисуем на 5332А порты для стыков с серверами
                for (int intCurrentServer = 0; intCurrentServer <= intTotalLogServers; intCurrentServer++)
                {
                    //Добавляем в спецификацию
                    list_Specification.Add("NR-SFP-10G-AOC 10G SFP+ <=> SFP+");

                    listShapesMesLogUpLinkPorts.Add(page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.4, doubNextPortStartPointY + 0.2));
                    listShapesMesLogUpLinkPorts[listShapesMesLogUpLinkPorts.Count - 1].Text = "XGE-" + (30 - intCurrentServer);
                    listShapesMesLogUpLinkPorts[listShapesMesLogUpLinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listShapesMesLogUpLinkPorts[listShapesMesLogUpLinkPorts.Count - 1].Data1 = "MES5332A";
                    listShapesMesLogUpLinkPorts[listShapesMesLogUpLinkPorts.Count - 1].Data2 = "XGE-" + (30 - intCurrentServer);
                    strPortName = "XGE-" + (30 - intCurrentServer);

                    //Log Lines
                    listShapesMesLogUpLinkLines.Add(page1.DrawLine(doubNextPortStartPointX + 0.4, doubNextPortStartPointY + 0.1, doubNextPortStartPointX + 2, doubNextPortStartPointY + 0.1));
                    listShapesMesLogUpLinkCircles.Add(page1.DrawOval(doubNextPortStartPointX + 0.5 + 0.4 * (intCurrentServer % 2), doubNextPortStartPointY - 0.1, doubNextPortStartPointX + 0.9 + 0.4 * (intCurrentServer % 2), doubNextPortStartPointY + 0.3));
                    listShapesMesLogUpLinkCircles[listShapesMesLogUpLinkCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                    listShapesMesLogUpLinkCircles[listShapesMesLogUpLinkCircles.Count - 1].Text = listShapesMesLogDownLinkCircles.Count + listShapesMesLogUpLinkCircles.Count + " LG";
                    listShapesMesLogUpLinkFakeRects.Add(page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY + 0.1, doubNextPortStartPointX + 2, doubNextPortStartPointY + 0.1));

                    // MySQL
                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{strPortName}', 9, 36);";
                    command.ExecuteNonQuery();
                    listShapesMesLogUpLinkPorts[listShapesMesLogUpLinkPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                    doubNextPortStartPointY += 0.2;
                };



                //Рисуем на MES5332A порт GE-32 - на стык с континентом
                Visio.Shape visMes5332aToContinentPort = page1.DrawRectangle(doubStartPointNextShapeX + 3.6, doubStartPointNextShapeY + 4.3, doubStartPointNextShapeX + 4, doubStartPointNextShapeY + 4.5);
                visMes5332aToContinentPort.Text = "GE-32";
                visMes5332aToContinentPort.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                //IC Lines
                Visio.Shape visMes5332aToContinentLine = page1.DrawLine(doubStartPointNextShapeX + 2, doubStartPointNextShapeY + 4.4, doubStartPointNextShapeX + 3.6, doubStartPointNextShapeY + 4.4);
                Visio.Shape visMes5332aToContinentCircle = page1.DrawOval(doubStartPointNextShapeX + 2.6, doubStartPointNextShapeY + 4.2, doubStartPointNextShapeX + 3, doubStartPointNextShapeY + 4.6);
                visMes5332aToContinentCircle.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                visMes5332aToContinentCircle.Text = listShapesMesLogDownLinkPorts.Count + listShapesMesLogUpLinkPorts.Count + 1 + " LG";
                Visio.Shape visMes5332aToContinentFakeRects = page1.DrawRectangle(doubStartPointNextShapeX + 2, doubStartPointNextShapeY + 4.4, doubStartPointNextShapeX + 2, doubStartPointNextShapeY + 4.4);
                visContinentToMes5332aCircle.Text = listShapesMesLogDownLinkPorts.Count + listShapesMesLogUpLinkPorts.Count + 1 + " LG";
                visMes5332aToContinentFakeRects.AutoConnect(visContinentToMes5332aFakeRects, Visio.VisAutoConnectDir.visAutoConnectDirNone);

                //Добавляем в КЖ запись "5332А - Континент" - сразу обе стороны.
                intGlobalCableCounter++;
                listCableJournal_Log.Add(new Dictionary<string, string>());
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Number", listShapesMesLogDownLinkPorts.Count + listShapesMesLogUpLinkPorts.Count + 1 + " LG");
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Name", "MES5332A --- " + strContinentHostname);
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Device_A", "MES5332A");
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Port_A", "GE-32");
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Device_B", strContinentHostname);
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Port_B", strContinentToMes5332aPort);
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Type", "LC-LC/UPC MM (50/125мкм) OM4 (3.0мм) 2м");
                // MySQL
                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'GE-32', 4, 1202);";
                command.ExecuteNonQuery();
                strNeighborPortId = visContinentToMes5332aPort.Data3;
                strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                command.ExecuteNonQuery();



                // Сервер СПХД

                //Console.WriteLine("Добавляем серверы.");

                double doubServersStartX = doubStartPointNextShapeX + 10 + 0.7 * listDeviceMgmtPorts.Count;
                double doubServersStartY = doubStartPointNextShapeY + 3;

                Visio.Shape shapeServerShdDevice = page1.DrawRectangle(doubServersStartX, doubServersStartY, doubServersStartX + 2, doubServersStartY + 0.6);

                strCurrentDeviceHostname = "Сервер СПХД";
                strNameForRackTables = $"{strFullSiteIndex}-SPHD-1";
                shapeServerShdDevice.Text = strCurrentDeviceHostname;
                shapeServerShdDevice.get_Cells("FillForegnd").FormulaU = "=RGB(210,180,140)";

                //Добавляем в спецификацию
                list_Specification.Add("СПХД");
                List<Visio.Shape> listShapesSrvPortsToLog = new List<Visio.Shape>();

                /////////////// Добавляем шасси в БД
                intCurrentRackSlot += 5;
                intDeviceId = FillRackSlot(command, strObjectIndex, intNetSrvRackId, intCurrentRackSlot, 4, strNameForRackTables);

                listShapesSrvPortsToLog.Add(page1.DrawRectangle(doubServersStartX - 0.5, doubStartPointNextShapeY + 3.1, doubServersStartX, doubStartPointNextShapeY + 3.3));
                listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].Text = "XGE-1 (верхний)";
                listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                //Log Lines
                listDeviceLogLines.Add(page1.DrawLine(doubServersStartX - 2, doubStartPointNextShapeY + 3.2, doubServersStartX - 0.5, doubStartPointNextShapeY + 3.2));
                listDeviceLogCircles.Add(page1.DrawOval(doubServersStartX - 1.1, doubStartPointNextShapeY + 3, doubServersStartX - 0.7, doubStartPointNextShapeY + 3.4));
                listDeviceLogCircles[listDeviceLogCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                listDeviceLogCircles[listDeviceLogCircles.Count - 1].Text = listDeviceLogCircles.Count + " LG";
                listDeviceLogFakeRects.Add(page1.DrawRectangle(doubServersStartX - 2, doubStartPointNextShapeY + 3.2, doubServersStartX - 2, doubStartPointNextShapeY + 3.2));
                listDeviceLogFakeRects[listDeviceLogFakeRects.Count - 1].AutoConnect(listShapesMesLogUpLinkFakeRects[0], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                // MySQL
                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'XGE-1 (верхний)', 9, 36);";
                command.ExecuteNonQuery();
                strNeighborPortId = listShapesMesLogUpLinkPorts[0].Data3;
                strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                command.ExecuteNonQuery();

                //КЖ
                intGlobalCableCounter++;
                listCableJournal_Log.Add(new Dictionary<string, string>());
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Number", listDeviceLogPorts.Count + 1 + " LG");
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Name", "СПХД --- " + listShapesMesLogUpLinkPorts[0].Data1);
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Device_A", strCurrentDeviceHostname);
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Port_A", "XGE-1 (верхний)");
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Device_B", listShapesMesLogUpLinkPorts[0].Data1);
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Port_B", listShapesMesLogUpLinkPorts[0].Data2);
                listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Type", "NR-SFP-10G-AOC 10G SFP+ <=> SFP+ (10G, SFP+, AOC, 2м)");

                listDeviceMgmtPorts.Add(page1.DrawRectangle(doubServersStartX + 2, doubStartPointNextShapeY + 3.1, doubServersStartX + 2.5, doubStartPointNextShapeY + 3.3));
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "GE-4";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "GE-4";
                // MySQL Mgmt.
                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'GE-4', 1, 24);";
                command.ExecuteNonQuery();
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                //Mgmt Lines
                listDeviceMgmtLines.Add(page1.DrawLine(doubServersStartX + 2.5, doubStartPointNextShapeY + 3.2, doubServersStartX + 5, doubStartPointNextShapeY + 3.2));
                listDeviceMgmtCircles.Add(page1.DrawOval(doubServersStartX + 2.8, doubStartPointNextShapeY + 3, doubServersStartX + 3.2, doubStartPointNextShapeY + 3.4));
                listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubServersStartX + 5, doubStartPointNextShapeY + 3.2, doubServersStartX + 5, doubStartPointNextShapeY + 3.2));
                listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "GE-4";

                listDeviceMgmtPorts.Add(page1.DrawRectangle(doubServersStartX + 2, doubStartPointNextShapeY + 3.3, doubServersStartX + 2.5, doubStartPointNextShapeY + 3.5));
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "Mgmt";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "Mgmt";
                //Mgmt Lines
                listDeviceMgmtLines.Add(page1.DrawLine(doubServersStartX + 2.5, doubStartPointNextShapeY + 3.4, doubServersStartX + 5.2, doubStartPointNextShapeY + 3.4));
                listDeviceMgmtCircles.Add(page1.DrawOval(doubServersStartX + 3.2, doubStartPointNextShapeY + 3.2, doubServersStartX + 3.6, doubStartPointNextShapeY + 3.6));
                listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubServersStartX + 5.2, doubStartPointNextShapeY + 3.4, doubServersStartX + 5.2, doubStartPointNextShapeY + 3.4));
                listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "Mgmt";
                // MySQL Mgmt.
                command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'Mgmt', 1, 24);";
                command.ExecuteNonQuery();
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                vsoWindow.DeselectAll();
                vsoWindow.Select(shapeServerShdDevice, 2);
                vsoWindow.Select(listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1], 2);
                vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1], 2);
                vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 2], 2);
                vsoSelection = vsoWindow.Selection;
                vsoSelection.Group();


                //Серверы СПФС

                int intCurrentLogUplinkPort = 0;

                List<Visio.Shape> listShapesSpfsDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesSpfsPorts = new List<Visio.Shape>();

                double doubSfsdDeviceX;
                double doubSfsdDeviceY = doubStartPointNextShapeY + 4;

                for (int intCurrentSpfsDevice = 1; intCurrentSpfsDevice <= intTotalLogServers; intCurrentSpfsDevice++)
                {
                    doubSfsdDeviceX = doubStartPointNextShapeX + 10 + 0.7 * listDeviceMgmtPorts.Count + 0.1 * listShapesSpfsDevices.Count;
                    intCurrentLogUplinkPort++;
                    listShapesSpfsDevices.Add(page1.DrawRectangle(doubSfsdDeviceX, doubSfsdDeviceY, doubSfsdDeviceX + 2, doubSfsdDeviceY + 0.6));
                    strCurrentDeviceHostname = "Сервер СПФС (" + intCurrentSpfsDevice + ")";
                    strNameForRackTables = $"{strFullSiteIndex}-SPFS-{intCurrentSpfsDevice}";
                    listShapesSpfsDevices[listShapesSpfsDevices.Count - 1].Text = strCurrentDeviceHostname;
                    listShapesSpfsDevices[listShapesSpfsDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(210,180,140)";
                    list_Specification.Add("СПФС");
                    listShapesSrvPortsToLog.Add(page1.DrawRectangle(doubSfsdDeviceX - 0.5, doubSfsdDeviceY + 0.1, doubSfsdDeviceX, doubSfsdDeviceY + 0.3));
                    listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].Text = "XGE-1 (верхний)";
                    listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    //Log Lines
                    listDeviceLogLines.Add(page1.DrawLine(doubSfsdDeviceX - 5, doubSfsdDeviceY + 0.2, doubSfsdDeviceX - 0.5, doubSfsdDeviceY + 0.2));
                    listDeviceLogCircles.Add(page1.DrawOval(doubSfsdDeviceX - 1.1, doubSfsdDeviceY, doubSfsdDeviceX - 0.7, doubSfsdDeviceY + 0.4));
                    listDeviceLogCircles[listDeviceLogCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                    listDeviceLogCircles[listDeviceLogCircles.Count - 1].Text = listDeviceLogPorts.Count + 1 + intCurrentSpfsDevice + " LG";
                    listDeviceLogFakeRects.Add(page1.DrawRectangle(doubSfsdDeviceX - 5, doubSfsdDeviceY + 0.2, doubSfsdDeviceX - 5, doubSfsdDeviceY + 0.2));
                    listDeviceLogFakeRects[listDeviceLogFakeRects.Count - 1].AutoConnect(listShapesMesLogUpLinkFakeRects[intCurrentLogUplinkPort], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                    /////////////// Добавляем шасси в БД
                    intCurrentRackSlot += 1;
                    intDeviceId = FillRackSlot(command, strObjectIndex, intNetSrvRackId, intCurrentRackSlot, 4, strNameForRackTables);

                    // MySQL
                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'XGE-1 (верхний)', 9, 36);";
                    command.ExecuteNonQuery();
                    strNeighborPortId = listShapesMesLogUpLinkPorts[intCurrentLogUplinkPort].Data3;
                    strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                    command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                    command.ExecuteNonQuery();

                    //КЖ
                    intGlobalCableCounter++;
                    listCableJournal_Log.Add(new Dictionary<string, string>());
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Number", listDeviceLogPorts.Count + 1 + intCurrentSpfsDevice + " LG");
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Name", strCurrentDeviceHostname + " --- " + listShapesMesLogUpLinkPorts[listShapesSrvPortsToLog.Count - 1].Data1);
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Device_A", strCurrentDeviceHostname);
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Port_A", "XGE-1 (верхний)");
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Device_B", listShapesMesLogUpLinkPorts[listShapesSrvPortsToLog.Count - 1].Data1);
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Port_B", listShapesMesLogUpLinkPorts[listShapesSrvPortsToLog.Count - 1].Data2);
                    listCableJournal_Log[listCableJournal_Log.Count - 1].Add("Cable_Type", "NR-SFP-10G-AOC 10G SFP+ <=> SFP+ (10G, SFP+, AOC, 2м)");

                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubSfsdDeviceX + 2, doubSfsdDeviceY + 0.1, doubSfsdDeviceX + 2.5, doubSfsdDeviceY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "GE-4";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "GE-4";
                    //Mgmt Lines
                    listDeviceMgmtLines.Add(page1.DrawLine(doubSfsdDeviceX + 2.5, doubSfsdDeviceY + 0.2, doubSfsdDeviceX + 5, doubSfsdDeviceY + 0.2));
                    listDeviceMgmtCircles.Add(page1.DrawOval(doubSfsdDeviceX + 2.8, doubSfsdDeviceY, doubSfsdDeviceX + 3.2, doubSfsdDeviceY + 0.4));
                    listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                    listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                    listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubSfsdDeviceX + 5, doubSfsdDeviceY + 0.2, doubSfsdDeviceX + 5, doubSfsdDeviceY + 0.2));
                    listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "GE-4";
                    // MySQL Mgmt.
                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'GE-4', 1, 24);";
                    command.ExecuteNonQuery();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubSfsdDeviceX + 2, doubSfsdDeviceY + 0.3, doubSfsdDeviceX + 2.5, doubSfsdDeviceY + 0.5));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "Mgmt";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "Mgmt";
                    //Mgmt Lines
                    listDeviceMgmtLines.Add(page1.DrawLine(doubSfsdDeviceX + 2.5, doubSfsdDeviceY + 0.4, doubSfsdDeviceX + 5.2, doubSfsdDeviceY + 0.4));
                    listDeviceMgmtCircles.Add(page1.DrawOval(doubSfsdDeviceX + 3.2, doubSfsdDeviceY + 0.2, doubSfsdDeviceX + 3.6, doubSfsdDeviceY + 0.6));
                    listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                    listDeviceMgmtCircles[listDeviceMgmtCircles.Count - 1].Text = Convert.ToString(listDeviceMgmtCircles.Count) + " М";
                    listDeviceMgmtFakeRects.Add(page1.DrawRectangle(doubSfsdDeviceX + 5.2, doubSfsdDeviceY + 0.4, doubSfsdDeviceX + 5.2, doubSfsdDeviceY + 0.4));
                    listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtFakeRects[listDeviceMgmtFakeRects.Count - 1].Data2 = "Mgmt";
                    // MySQL Mgmt.
                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'Mgmt', 1, 24);";
                    command.ExecuteNonQuery();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(LastUsedId(command, "Port"));

                    vsoWindow.DeselectAll();
                    vsoWindow.Select(listShapesSpfsDevices[listShapesSpfsDevices.Count - 1], 2);
                    vsoWindow.Select(listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1], 2);
                    vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1], 2);
                    vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 2], 2);
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();

                    doubSfsdDeviceY += 0.8;
                }

                // Рисуем 3348.
                int intMes3348ChassisNumber = CalculateDevicesQuantity(listDeviceMgmtPorts.Count, 46);
                //Console.WriteLine($"Портов MGMT: {listDeviceMgmtPorts.Count}");
                doubStartPointNextShapeX = doubNextPortStartPointX + 3.5;
                string strPrevious3348Hostname = "";
                string strPrevious3348Portname = "";
                strCurrentDeviceHostname = "";
                int[] arrDeviceIdOnMes = new int[10];
                List<Visio.Shape> listShapesMesMgmtDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesMgmtPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesMgmtLines = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesMgmtFakeRects = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesMgmtCircles = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesUplinkPorts = new List<Visio.Shape>();
                Visio.Shape[,] arrMes3348UplinkPorts = new Visio.Shape[10, 2];
                Visio.Shape[,] arrMes3348UplinkLines = new Visio.Shape[10, 2];
                Visio.Shape[,] arrMes3348UplinkCircles = new Visio.Shape[10, 2];
                Visio.Shape[,] arrMes3348UplinkFakeRects = new Visio.Shape[10, 2];
                int intCurrentMgmtPortCounter = 0;
                int intMes48Port = 46;
                int intMess3348CurrentChassis = 0;
                double doubSwitchWidth;
                intCurrentRackSlot += 5;
                for (int intCurrentMesPort = 1; intCurrentMesPort <= listDeviceMgmtPorts.Count; intCurrentMesPort++)    // Не считать порты 47 и 48 в общем списке.
                {
                    // Рисуем новое шасси, если текущий порт = 1, 47 и т.д.
                    if (intCurrentMesPort % 46 == 1)
                    //////////////////////  Add New MES Switch  //////////////////////////
                    {
                        if (listDeviceMgmtPorts.Count - intCurrentMesPort <= 46) doubSwitchWidth = (listDeviceMgmtPorts.Count - intCurrentMesPort) * 0.22 + 1;
                        else doubSwitchWidth = 46 * 0.22;

                        doubNextPortStartPointX += 1;
                        listShapesMesMgmtDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + doubSwitchWidth, doubStartPointNextShapeY + 1));
                        strCurrentDeviceHostname = "MES3348 (" + listShapesMesMgmtDevices.Count + ")";
                        strNameForRackTables = $"{strFullSiteIndex}-MGMT{listShapesMesMgmtDevices.Count}";
                        listShapesMesMgmtDevices[listShapesMesMgmtDevices.Count - 1].Text = strCurrentDeviceHostname;
                        listShapesMesMgmtDevices[listShapesMesMgmtDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(176,196,222)";
                        listShapesMesMgmtDevices[listShapesMesMgmtDevices.Count - 1].Data3 = strCurrentDeviceHostname;
                        intMes48Port = 0;
                        doubNextPortStartPointX = doubStartPointNextShapeX + 0.1;
                        doubNextPortStartPointY = doubStartPointNextShapeY - 0.3;

                        //MySQL.

                        intCurrentRackSlot += 1;
                        intDeviceId = FillRackSlot(command, strObjectIndex, intNetSrvRackId, intCurrentRackSlot, 8, strNameForRackTables);
                        //arrDeviceIdOnMes

                        //На новом свитче рисуем три служебных порта
                        //Порт GE-48 вставляется слева на любом свитче.
                        intMess3348CurrentChassis++;
                        arrDeviceIdOnMes[intMess3348CurrentChassis] = intDeviceId;
                        arrMes3348UplinkPorts[intMess3348CurrentChassis, 0] = page1.DrawRectangle(doubStartPointNextShapeX - 0.4, doubStartPointNextShapeY + 0.2, doubStartPointNextShapeX, doubStartPointNextShapeY + 0.4);
                        arrMes3348UplinkPorts[intMess3348CurrentChassis, 0].Text = "GE-48";
                        arrMes3348UplinkPorts[intMess3348CurrentChassis, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        //IC Lines
                        arrMes3348UplinkLines[intMess3348CurrentChassis, 0] = page1.DrawLine(doubStartPointNextShapeX - 1.8, doubStartPointNextShapeY + 0.3, doubStartPointNextShapeX - 0.4, doubStartPointNextShapeY + 0.3);
                        arrMes3348UplinkCircles[intMess3348CurrentChassis, 0] = page1.DrawOval(doubStartPointNextShapeX - 0.9, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX - 0.5, doubStartPointNextShapeY + 0.5);
                        arrMes3348UplinkCircles[intMess3348CurrentChassis, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                        arrMes3348UplinkCircles[intMess3348CurrentChassis, 0].Text = listDeviceMgmtFakeRects.Count + 1 + intMess3348CurrentChassis + " M";
                        arrMes3348UplinkFakeRects[intMess3348CurrentChassis, 0] = page1.DrawRectangle(doubStartPointNextShapeX - 1.8, doubStartPointNextShapeY + 0.3, doubStartPointNextShapeX - 1.8, doubStartPointNextShapeY + 0.3);
                        //Если свитч первый, его порт GE-48 стыкуется с континентом.
                        if (intMess3348CurrentChassis == 1)
                        {
                            arrMes3348UplinkFakeRects[intMess3348CurrentChassis, 0].AutoConnect(visContinentToMes3348FakeRect, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            visContinentToMes3348Circle.Text = listDeviceMgmtFakeRects.Count + 1 + intMess3348CurrentChassis + " M";
                            strPrevious3348Hostname = strContinentHostname;
                            strPrevious3348Portname = strContinentToMes3348Port;

                            intGlobalCableCounter++;
                            listCableJournal_Management.Add(new Dictionary<string, string>());
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", listDeviceMgmtFakeRects.Count + 2 + " M");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", strContinentHostname + " - " + strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", strContinentHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", strContinentToMes3348Port);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", "GE-48");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "UTP cat. 5e (RJ45-RJ45)");
                            doubContStartY += 0.2;

                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'GE-48', 1, 24);";
                            command.ExecuteNonQuery();
                            strNeighborPortId = visContinentToMes3348Port.Data3;
                            strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                            command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                            command.ExecuteNonQuery();

                        }
                        //Если свитч не первый, его порт GE-48 стыкуется с GE-47 предыдущего свитча.
                        else
                        {
                            arrMes3348UplinkFakeRects[intMess3348CurrentChassis, 0].AutoConnect(arrMes3348UplinkFakeRects[intMess3348CurrentChassis - 1, 1], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            arrMes3348UplinkCircles[intMess3348CurrentChassis - 1, 1].Text = listDeviceMgmtFakeRects.Count + 1 + intMess3348CurrentChassis + " M";
                            //удалить после свопа
                            strPrevious3348Portname = "GE-47";
                            intGlobalCableCounter++;
                            listCableJournal_Management.Add(new Dictionary<string, string>());
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", listDeviceMgmtFakeRects.Count + 1 + intMess3348CurrentChassis + " M");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", strPrevious3348Hostname + " --- " + strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", strPrevious3348Hostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", "GE-47");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", "GE-48");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "UTP cat. 5e (RJ45-RJ45)");
                            strPrevious3348Hostname = strContinentHostname;
                            strPrevious3348Portname = strContinentToMes3348Port;

                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'GE-48', 1, 24);";
                            command.ExecuteNonQuery();
                            strNeighborPortId = Convert.ToString(intLast3348Port);
                            strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                            command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                            command.ExecuteNonQuery();
                        };
                        //Если свитч последний, его порт GE-47 рисуется слева, стыковка с ООВ MES5332A.
                        if (intMess3348CurrentChassis == intMes3348ChassisNumber)
                        {
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1] = page1.DrawRectangle(doubStartPointNextShapeX - 0.4, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX, doubStartPointNextShapeY + 0.6);
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1].Text = "GE-47";
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            //IC Lines
                            arrMes3348UplinkLines[intMess3348CurrentChassis, 1] = page1.DrawLine(doubStartPointNextShapeX - 1.6, doubStartPointNextShapeY + 0.5, doubStartPointNextShapeX - 0.4, doubStartPointNextShapeY + 0.5);
                            arrMes3348UplinkCircles[intMess3348CurrentChassis, 1] = page1.DrawOval(doubStartPointNextShapeX - 1.4, doubStartPointNextShapeY + 0.3, doubStartPointNextShapeX - 1, doubStartPointNextShapeY + 0.7);
                            arrMes3348UplinkCircles[intMess3348CurrentChassis, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrMes3348UplinkCircles[intMess3348CurrentChassis, 1].Text = listDeviceMgmtFakeRects.Count + 1 + " M";
                            visMes5332aTo3348Circle.Text = listDeviceMgmtFakeRects.Count + 1 + " M";
                            arrMes3348UplinkFakeRects[intMess3348CurrentChassis, 1] = page1.DrawRectangle(doubStartPointNextShapeX - 1.6, doubStartPointNextShapeY + 0.5, doubStartPointNextShapeX - 1.6, doubStartPointNextShapeY + 0.5);
                            arrMes3348UplinkFakeRects[intMess3348CurrentChassis, 1].AutoConnect(visMes5332aTo3348FakeRect, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            intGlobalCableCounter++;
                            listCableJournal_Management.Add(new Dictionary<string, string>());
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", listDeviceMgmtFakeRects.Count + 1 + " M");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", "MES5332A --- " + strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", "MES5332A");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", "OOB");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", "GE-47");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "UTP cat. 5e (RJ45-RJ45)");

                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'GE-47', 1, 24);";
                            command.ExecuteNonQuery();
                            strNeighborPortId = visMes5332aTo3348Port.Data3;
                            strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                            command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                            command.ExecuteNonQuery();
                        }
                        //Если свитч не последний, его порт GE-47 рисуется справа. Стык - при настройке следующего свитча.
                        else
                        {
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1] = page1.DrawRectangle(doubStartPointNextShapeX + doubSwitchWidth, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + doubSwitchWidth + 0.4, doubStartPointNextShapeY + 0.6);
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1].Text = "GE-47";
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08"; ;
                            //IC Lines
                            arrMes3348UplinkLines[intMess3348CurrentChassis, 1] = page1.DrawLine(doubStartPointNextShapeX + doubSwitchWidth + 0.4, doubStartPointNextShapeY + 0.5, doubStartPointNextShapeX + doubSwitchWidth + 1.4, doubStartPointNextShapeY + 0.5);
                            arrMes3348UplinkCircles[intMess3348CurrentChassis, 1] = page1.DrawOval(doubStartPointNextShapeX + doubSwitchWidth + 0.5, doubStartPointNextShapeY + 0.3, doubStartPointNextShapeX + doubSwitchWidth + 0.9, doubStartPointNextShapeY + 0.7);
                            arrMes3348UplinkCircles[intMess3348CurrentChassis, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                            arrMes3348UplinkFakeRects[intMess3348CurrentChassis, 1] = page1.DrawRectangle(doubStartPointNextShapeX + doubSwitchWidth + 1.4, doubStartPointNextShapeY + 0.5, doubStartPointNextShapeX + doubSwitchWidth + 1.4, doubStartPointNextShapeY + 0.5);
                            strPrevious3348Hostname = strCurrentDeviceHostname;
                            // MySQL
                            command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, 'GE-47', 1, 24);";
                            command.ExecuteNonQuery();
                            intLast3348Port = LastUsedId(command, "Port");
                        };
                        doubStartPointNextShapeX += 15;
                    };
                    // Рисуем Downlink-порты MES3348.
                    intMes48Port++;
                    listShapesMesMgmtPorts.Add(page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.4, doubNextPortStartPointY + 0.2));
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Text = "GE-" + intMes48Port;
                    doubNextPortStartPointX += 0.2;
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Rotate90();
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Data1 = "MES3348 (" + listShapesMesMgmtDevices.Count + ")";
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Data2 = "GE-" + intMes48Port;
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Data3 = strCurrentDeviceHostname;
                    vsoWindow.Select(listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1], 2);
                    listShapesMesMgmtLines.Add(page1.DrawLine(doubNextPortStartPointX, doubNextPortStartPointY - 0.1, doubNextPortStartPointX, doubNextPortStartPointY - 1.4 - 0.1 * listShapesMesMgmtLines.Count));
                    listShapesMesMgmtFakeRects.Add(page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY - 1.4 - 0.1 * listShapesMesMgmtFakeRects.Count, doubNextPortStartPointX, doubNextPortStartPointY - 1.4 - 0.1 * listShapesMesMgmtFakeRects.Count));
                    listShapesMesMgmtCircles.Add(page1.DrawOval(doubNextPortStartPointX - 0.2, doubNextPortStartPointY - 0.7 - 0.4 * (intMes48Port % 2), doubNextPortStartPointX + 0.2, doubNextPortStartPointY - 0.3 - 0.4 * (intMes48Port % 2)));
                    listShapesMesMgmtCircles[listShapesMesMgmtCircles.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.16";
                    listShapesMesMgmtCircles[listShapesMesMgmtCircles.Count - 1].Text = listShapesMesMgmtPorts.Count + " М";


                };
                doubStartPointNextShapeX += listDeviceMgmtPorts.Count * 0.25 + 1;
                int intMesPortCurrentPort = 0;
                // Connect + КЖ + SQL.
                intMess3348CurrentChassis = 0;
                foreach (Visio.Shape objDeviceMgmtPort in listDeviceMgmtFakeRects)
                {
                    objDeviceMgmtPort.AutoConnect(listShapesMesMgmtFakeRects[intMesPortCurrentPort], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                    intGlobalCableCounter++;
                    intCurrentMgmtPortCounter++;
                    listCableJournal_Management.Add(new Dictionary<string, string>());
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", intCurrentMgmtPortCounter + " М");
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", objDeviceMgmtPort.Data1 + " --- " + listShapesMesMgmtPorts[intMesPortCurrentPort].Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", objDeviceMgmtPort.Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", objDeviceMgmtPort.Data2);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", listShapesMesMgmtPorts[intMesPortCurrentPort].Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", listShapesMesMgmtPorts[intMesPortCurrentPort].Data2);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "UTP cat. 5e (RJ45-RJ45)");

                    // MySQL
                    //int[] arrDeviceIdOnMes = new int[10];
                    if (listShapesMesMgmtPorts[intMesPortCurrentPort].Data2 == "GE-1") intMess3348CurrentChassis++;
                    //Console.WriteLine($"MES Number: {intDeviceId}, PortName: {listShapesMesMgmtPorts[intMesPortCurrentPort].Data2}");
                    intDeviceId = arrDeviceIdOnMes[intMess3348CurrentChassis];
                    command.CommandText = $"insert into Port (object_id, name, iif_id, type) values ({intDeviceId}, '{listShapesMesMgmtPorts[intMesPortCurrentPort].Data2}', 1, 24);";
                    command.ExecuteNonQuery();
                    //Console.WriteLine($"Чек");
                    strNeighborPortId = listDeviceMgmtPorts[intMesPortCurrentPort].Data3;
                    strLocalPortId = Convert.ToString(LastUsedId(command, "Port"));
                    command.CommandText = $"insert into Link (porta, portb) values ({strNeighborPortId}, {strLocalPortId});";
                    command.ExecuteNonQuery();

                    intMesPortCurrentPort++;
                };
                intMesPortCurrentPort = 0;

                //////  Group 5332A Ports ////////
                vsoWindow.DeselectAll();
                vsoWindow.Select(shapeMesLogDevice, 2);
                foreach (Visio.Shape objMesLogPort in listShapesMesLogUpLinkPorts)
                {
                    vsoWindow.Select(objMesLogPort, 2);
                };
                foreach (Visio.Shape objMesLogPort in listShapesMesLogDownLinkPorts)
                {
                    vsoWindow.Select(objMesLogPort, 2);
                };
                vsoSelection = vsoWindow.Selection;
                vsoSelection.Group();

                doc.SaveAs(strVsdFilePath);


                //Quit
                doc.Close();
                app.Quit();


                int intForeachIterator = 0;

                /////////////////////  Создание файла КЖ  ///////////////////////////////

                // Создаём файл КЖ и заполняем названия заголовков.
                Excel.Application xlApp2 = new Excel.Application();
                Excel.Workbook xlWorkbook2 = xlApp2.Workbooks.Add(Type.Missing);
                Excel.Worksheet xlWorksheet31 = (Excel.Worksheet)xlWorkbook2.Worksheets.get_Item(1);
                xlWorksheet31.Name = "Сводный КЖ";
                xlWorksheet31.Range[xlWorksheet31.Cells[1, 1], xlWorksheet31.Cells[1, 11]].Merge();
                xlWorksheet31.Cells[1, 1] = "Кабельный Журнал " + strObjectAddress + "       (Сводный)";
                Excel.Range formatObjectAddress31;
                formatObjectAddress31 = xlWorksheet31.get_Range("a1", "a1");
                formatObjectAddress31.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSeaGreen);
                xlWorksheet31.Cells[2, 1] = "Номер кабельного соединения";
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 1], xlWorksheet31.Cells[2, 1]].WrapText = true;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 1], xlWorksheet31.Cells[3, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 1], xlWorksheet31.Cells[3, 1]].Merge();
                xlWorksheet31.Cells[2, 2] = "Наименование участка";
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 2], xlWorksheet31.Cells[3, 2]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 2], xlWorksheet31.Cells[3, 2]].Merge();
                xlWorksheet31.Cells[2, 3] = "Откуда";
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 3], xlWorksheet31.Cells[2, 4]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 3], xlWorksheet31.Cells[2, 4]].Merge();
                xlWorksheet31.Cells[3, 3] = "№№ стойки, шкафа;\nнаименование оборудования";
                xlWorksheet31.Cells[3, 4] = "Плата (слот) / гнездо (порт)";
                xlWorksheet31.Cells[2, 5] = "Куда";
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 5], xlWorksheet31.Cells[2, 6]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 5], xlWorksheet31.Cells[2, 6]].Merge();
                xlWorksheet31.Cells[3, 5] = "№№ стойки, шкафа;\nнаименование оборудования";
                xlWorksheet31.Cells[3, 6] = "Плата (слот) / гнездо (порт)";
                xlWorksheet31.Cells[2, 7] = "Марка, ёмкость кабеля";
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 7], xlWorksheet31.Cells[3, 7]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 7], xlWorksheet31.Cells[3, 7]].Merge();
                xlWorksheet31.Cells[2, 8] = "Количество кусков (шт)";
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 8], xlWorksheet31.Cells[3, 8]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 8], xlWorksheet31.Cells[3, 8]].Merge();
                xlWorksheet31.Cells[2, 9] = "Длина куска (м)";
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 9], xlWorksheet31.Cells[3, 9]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 9], xlWorksheet31.Cells[3, 9]].Merge();
                xlWorksheet31.Cells[2, 10] = "Общая длина (м)";
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 10], xlWorksheet31.Cells[3, 10]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 10], xlWorksheet31.Cells[3, 10]].Merge();
                xlWorksheet31.Cells[2, 11] = "Примечания";
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 11], xlWorksheet31.Cells[3, 11]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 11], xlWorksheet31.Cells[3, 11]].Merge();

                // Форматируем ячейки заголовков.
                Excel.Range formatHeaders;
                formatHeaders = xlWorksheet31.get_Range("a2", "k3");
                formatHeaders.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.BurlyWood);
                Excel.Range formatRange;
                Excel.Borders border;
                formatRange = xlWorksheet31.get_Range("a1").EntireRow.EntireColumn;
                formatRange.NumberFormat = "@";

                // LAN - Байпасы
                xlWorksheet31.Range[xlWorksheet31.Cells[4, 1], xlWorksheet31.Cells[4, 11]].Merge();
                xlWorksheet31.Range[xlWorksheet31.Cells[4, 1], xlWorksheet31.Cells[4, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                formatObjectAddress31 = xlWorksheet31.get_Range("a4", "a4");
                formatObjectAddress31.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);
                xlWorksheet31.Cells[4, 1] = "Сегмент LAN";
                intJournalCurrentRow = 4;
                foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_LAN_Bypass)
                {
                    intJournalCurrentRow++;
                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"]; ;
                    xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Side_A_Name"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Side_A_Port"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Side_B_Name"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Side_B_Port"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 11] = dictCableRecord["Comment"];
                };

                // WAN - Байпасы
                intJournalCurrentRow++;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);
                xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Сегмент WAN";
                foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_WAN_Bypass)
                {
                    intJournalCurrentRow++;
                    xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Side_B_Name"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Side_B_Port"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"]; ;
                    xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Side_A_Name"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Side_A_Port"];

                    xlWorksheet31.Cells[intJournalCurrentRow, 11] = dictCableRecord["Comment"];
                };

                // Байпасы - Что-то
                intJournalCurrentRow++;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);
                // Байпасы - Фильтр
                if (boolNoBalancer)
                {
                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Байпасы - Фильтр";
                    foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_Bypass_Balancer)
                    {
                        intJournalCurrentRow++;
                        xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"]; ;
                        xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Device_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Port_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Device_B_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Port_B_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 7] = dictCableRecord["Cable_Type"];
                    };
                }
                // Байпасы - Балансеры
                else
                {
                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Байпасы - Балансировщики";
                    intForeachIterator = 0;
                    foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_Bypass_Balancer)
                    {
                        intJournalCurrentRow++;

                        if (intLinkCounter40 == 0 && intLinkCounter100 == 0)
                        {
                            intForeachIterator++;
                            if ((intForeachIterator + 3) % 4 == 0)
                            {
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow + 3, 1]].Merge();
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 2], xlWorksheet31.Cells[intJournalCurrentRow + 3, 2]].Merge();
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 5], xlWorksheet31.Cells[intJournalCurrentRow + 3, 5]].Merge();
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow + 3, 7]].Merge();
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 8], xlWorksheet31.Cells[intJournalCurrentRow + 3, 8]].Merge();
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 9], xlWorksheet31.Cells[intJournalCurrentRow + 3, 9]].Merge();
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 10], xlWorksheet31.Cells[intJournalCurrentRow + 3, 10]].Merge();
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 11], xlWorksheet31.Cells[intJournalCurrentRow + 3, 11]].Merge();
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 2], xlWorksheet31.Cells[intJournalCurrentRow, 2]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 5], xlWorksheet31.Cells[intJournalCurrentRow, 5]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow, 7]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 8], xlWorksheet31.Cells[intJournalCurrentRow, 8]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 9], xlWorksheet31.Cells[intJournalCurrentRow, 9]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 10], xlWorksheet31.Cells[intJournalCurrentRow, 10]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 11], xlWorksheet31.Cells[intJournalCurrentRow, 11]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                                if (!boolEshelon)
                                {
                                    xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 3], xlWorksheet31.Cells[intJournalCurrentRow + 3, 3]].Merge();
                                    xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 3], xlWorksheet31.Cells[intJournalCurrentRow, 3]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                };
                            };
                        };
                        xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Device_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Port_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Device_B_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Port_B_Name"];
                        xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow, 7]].NumberFormat = "General";
                        xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow, 7]].Formula = $"=\"{dictCableRecord["Cable_Type"]}\" & I{intJournalCurrentRow}";
                        xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow, 7]].Calculate();
                    };



                    if (boolEshelon)
                    {
                        // Перемычка на хайвэях
                        intJournalCurrentRow++;
                        xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                        xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);
                        xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Перемычки Ecohighway";
                        foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_Highway_Peremychka)
                        {
                            intJournalCurrentRow++;
                            xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"];
                            xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                            xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Device_A_Name"];
                            xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Port_A_Name"];
                            xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Device_B_Name"];
                            xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Port_B_Name"];
                            xlWorksheet31.Cells[intJournalCurrentRow, 7] = dictCableRecord["Cable_Type"];
                        };
                    };


                    // Балансеры - Фильтры
                    intJournalCurrentRow++;
                    xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                    xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);
                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Балансировщики - Фильтры";
                    intForeachIterator = 0;
                    foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_Balancer_Filter)
                    {
                        intJournalCurrentRow++;
                        intForeachIterator++;
                        if ((intForeachIterator + 3) % 4 == 0)
                        {
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow + 3, 1]].Merge();
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 2], xlWorksheet31.Cells[intJournalCurrentRow + 3, 2]].Merge();
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 3], xlWorksheet31.Cells[intJournalCurrentRow + 3, 3]].Merge();
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 5], xlWorksheet31.Cells[intJournalCurrentRow + 3, 5]].Merge();
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow + 3, 7]].Merge();
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 8], xlWorksheet31.Cells[intJournalCurrentRow + 3, 8]].Merge();
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 9], xlWorksheet31.Cells[intJournalCurrentRow + 3, 9]].Merge();
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 10], xlWorksheet31.Cells[intJournalCurrentRow + 3, 10]].Merge();
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 11], xlWorksheet31.Cells[intJournalCurrentRow + 3, 11]].Merge();
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 2], xlWorksheet31.Cells[intJournalCurrentRow, 2]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 3], xlWorksheet31.Cells[intJournalCurrentRow, 3]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 5], xlWorksheet31.Cells[intJournalCurrentRow, 5]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow, 7]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 8], xlWorksheet31.Cells[intJournalCurrentRow, 8]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 9], xlWorksheet31.Cells[intJournalCurrentRow, 9]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 10], xlWorksheet31.Cells[intJournalCurrentRow, 10]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 11], xlWorksheet31.Cells[intJournalCurrentRow, 11]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        };
                        xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Device_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Device_B_Name"];
                        xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow, 7]].NumberFormat = "General";
                        xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow, 7]].Formula = $"=\"NR-QSFP-4X10G-AOC 40G QSFP\" & I{intJournalCurrentRow}";
                        xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 7], xlWorksheet31.Cells[intJournalCurrentRow, 7]].Calculate();
                        xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Port_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Port_B_Name"];
                    };
                };

                // Управление.
                intJournalCurrentRow++;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);
                xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Сегмент управления";
                foreach (Dictionary<string, string> dictCableRecord in listCableJournal_Management)
                {
                    intJournalCurrentRow++;
                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Device_A"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Port_A"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Device_B"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Port_B"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 7] = dictCableRecord["Cable_Type"];
                };

                // Логирование.
                intJournalCurrentRow++;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);
                xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Сегмент логирования";
                foreach (Dictionary<string, string> dictCableRecord in listCableJournal_Log)
                {
                    intJournalCurrentRow++;
                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Device_A"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Port_A"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Device_B"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Port_B"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 7] = dictCableRecord["Cable_Type"];
                };

                // Внешний стык.
                intJournalCurrentRow++;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);
                xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Внешний стык";
                intJournalCurrentRow++;
                xlWorksheet31.Cells[intJournalCurrentRow, 1] = "1 UL";
                xlWorksheet31.Cells[intJournalCurrentRow, 2] = "Канал управления (к " + strContinentHostname + ")";
                xlWorksheet31.Cells[intJournalCurrentRow, 3] = strContinentHostname;
                xlWorksheet31.Cells[intJournalCurrentRow, 4] = strContinentUplinkPort;
                //xlWorksheet31.Cells[intJournalCurrentRow, 5] = ;
                //xlWorksheet31.Cells[intJournalCurrentRow, 6] = ;
                //xlWorksheet31.Cells[intJournalCurrentRow, 7] = ;

                //Выравнивание ячеек
                xlWorksheet31.Range[xlWorksheet31.Cells[2, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                formatRange = xlWorksheet31.UsedRange;
                border = formatRange.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 2;
                formatRange.Columns.AutoFit();


                ////////////////////    Закрытие экселя ///////////////////////////////
                xlApp2.DisplayAlerts = false;
                xlWorkbook2.SaveAs(strExcelCablesFilePath);
                xlWorkbook2.Close();
                xlApp2.Quit();

                ///////////////////     Закрытие MySQL  ///////////////////////////////
                connection.Close();
                Console.WriteLine("Подключение закрыто.");

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkbook2);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp2);

                //Console.WriteLine("!!!!!!!!!!!!!! End !!!!!!!!!!!!!!!!");

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            }
            catch (Exception err)
            {
                Console.WriteLine($"Error: {err.Message}");
            }
            finally
            {
                Console.Write("\nPress any key to continue ...");
                Console.ReadKey();


            }
        }
    }
}