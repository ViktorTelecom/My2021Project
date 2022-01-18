using System;
using Visio = Microsoft.Office.Interop.Visio;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Generic;


namespace VeryFirstProject
{

    class Program
    {

        static int CalculateDevicesQuantity(int intLinkCounter, int intDivider)
        {
            if (intLinkCounter % intDivider == 0) return intLinkCounter / intDivider;
            else return intLinkCounter / intDivider + 1;
        }


        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        static void Main(string[] args)
        {

            //  Constants

            const double doubBandwidthOnFilter4160 = 60;                                                                       // 4160 Divider = 60 (New Constant)     



            //~~~~~~~~~~~~~~~~~~~   Start Try Block ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            try
            {
                //~~~~~~~~~~~~~~~~~~~~~~~   Declare File Path String Attributes ~~~~~~~~~~~~~~~~~~
                string strVsdFilePath;
                string strExcelCablesFilePath;
                string strXlsxFilePath = "";

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Open File Fialog Inside Thread (Open Reference Cable Journal) ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Thread thread1 = new Thread((ThreadStart)(() => {
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


                //~~~~~~~~~~~~  Calculate Other Files' Names

                strVsdFilePath = strXlsxFilePath.Replace(".xlsx", "-Layout.vsdx");                                               //~~~~~~~~~~~~  Calculate Other Files' Names
                strExcelCablesFilePath = strXlsxFilePath.Replace(".xlsx", "-Cables.xlsx");


                //~~~~~~~~~~~~~~    Open Reference Excel File
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(strXlsxFilePath);
                Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                Excel.Range xlRange = xlWorksheet.UsedRange;

                //Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                //Console.WriteLine($"Открыт шаблон: {strXlsxFilePath}");                                     //~~~~~~~~~~    Console Writeline (Excel File Path)

                int intTotalRows = xlRange.Rows.Count;

                Console.WriteLine($"Всего линков: {intTotalRows - 3}");                                                    //~~~~~~~~~~    Console Writeline (Total Links Number)





                //~~~~~~~~~~~~~~    Create Lists of Dictionaries    ~~~~~~~~~~~~~~~~~~~~~~~
                List<Dictionary<string, object>> listLanDevices = new List<Dictionary<string, object>>();           //LAN Devices
                List<Dictionary<string, object>> listWanDevices = new List<Dictionary<string, object>>();           //WAN Devices
                List<Dictionary<string, object>> listLanPorts = new List<Dictionary<string, object>>();             //LAN Ports
                List<Dictionary<string, object>> listWanPorts = new List<Dictionary<string, object>>();             //WAN Ports

                List<Dictionary<string, string>> listCableJournal_1 = new List<Dictionary<string, string>>();                   //Router-Bypass
                List<Dictionary<string, string>> list_CableJournal_Bypass_Filter = new List<Dictionary<string, string>>();      //Bypass-Balancer
                List<Dictionary<string, string>> listCableJournal_3 = new List<Dictionary<string, string>>();                   //Balancer-Filter
                List<Dictionary<string, string>> listCableJournal_4 = new List<Dictionary<string, string>>();                   //Для ТП
                List<Dictionary<string, string>> listCableJournal_Management = new List<Dictionary<string, string>>();          //Management&Log

                List<Dictionary<string, string>> list_CableJournal_LAN_Bypass = new List<Dictionary<string, string>>();              //Общий (LAN-Bypass)
                List<Dictionary<string, string>> list_CableJournal_WAN_Bypass = new List<Dictionary<string, string>>();              //Общий (WAN-Bypass)


                Dictionary<string, string>[,] arrCableJournal_LAN_Bypass = new Dictionary<string, string>[200, 200];
                Dictionary<string, string>[,] arrCableJournal_WAN_Bypass = new Dictionary<string, string>[200, 200];

                List<Dictionary<string, string>> list_CableJournal_Bypass_Balancer = new List<Dictionary<string, string>>();      //list_CableJournal_Bypass_Balancer
                List<Dictionary<string, string>> list_CableJournal_Balancer_Filter = new List<Dictionary<string, string>>();      //list_CableJournal_Balancer_Filter

                //~~~~~~~~~~~~~~    Create Lists of Visio Rectangles    ~~~~~~~~~~~~~~~~~~~~~~~
                List<Visio.Shape> listShapesLanDevices = new List<Visio.Shape>();                                    //LAN Devices Rects       
                List<Visio.Shape> listShapesWanDevices = new List<Visio.Shape>();                                    //WAN Devices Rects               
                List<Visio.Shape> listShapesLanPorts = new List<Visio.Shape>();                                      //LAN Ports Rects 
                List<Visio.Shape> listShapesWanPorts = new List<Visio.Shape>();                                      //LAN Ports Rects 


                Visio.Shape[] arrShapesLanPorts = new Visio.Shape[200];
                Visio.Shape[] arrShapesWanPorts = new Visio.Shape[200];


                Dictionary<string, string>[,,] arr_CableJournal_Bypass_Balancer = new Dictionary<string, string>[100, 200, 10];           //Bypass-Balancer
                Dictionary<string, string>[,,] arrCableJournal_Balancer_Filter = new Dictionary<string, string>[100, 200, 20];            //Balancer-Filter
                Dictionary<string, string>[] arrCableJournal_Bypass_Filter = new Dictionary<string, string>[200];            //Balancer-Filter


                //~~~~~~~~~~~~~~~~~ Read Common Data from File  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                double doubSummaryBandwidth = Convert.ToDouble(((Excel.Range)xlWorksheet.Cells[2, 19]).Value2.ToString());         // Read Summary BW from File [2,7]
                string strBypassModel = ((Excel.Range)xlWorksheet.Cells[2, 20]).Value2.ToString();                                 // Read Bypass Model     (Не было)
                string strFilterModel = ((Excel.Range)xlWorksheet.Cells[2, 21]).Value2.ToString();                                 // Read Filter Model         [2,19]
                string strContinentModel = ((Excel.Range)xlWorksheet.Cells[2, 22]).Value2.ToString();                              // Read Continent Model  (Не было)
                string strObjectAddress = ((Excel.Range)xlWorksheet.Cells[2, 16]).Value2.ToString();                               // Read Object Post Address  [2,16]
                //string strEolBypassIndication = ((Excel.Range)xlWorksheet.Cells[2, 24]).Value2.ToString();                       // Read EOL Bypass Y/N       [2,13]

                string strBalancerNumberFromInput = ((Excel.Range)xlWorksheet.Cells[2, 24]).Value2.ToString();                     // Read Balancers Number (не было)
                string strFilterNumberFromInput = ((Excel.Range)xlWorksheet.Cells[2, 25]).Value2.ToString();                       // Read Filter Number (не было)

                int intCurrentPortInChassis;

                int intStartBalancerPort;



                bool boolEolBypass = false;
                bool boolContinentIpcR300 = false;

                if (strBypassModel == "IBS1UP") boolEolBypass = true;
                if (strContinentModel == "300" || strContinentModel == "R300") boolContinentIpcR300 = true;

                //Console.WriteLine("~~~~~~~~~~~  CHECK   ~~~~~~~~~~~~~~~~");

                int intTotalFiltersFromBw = Convert.ToInt32(Math.Ceiling(doubSummaryBandwidth / doubBandwidthOnFilter4160));        //Calculated Total Filters Number (4120 or 4160)
                int intTotalLogServers = Convert.ToInt32(Math.Ceiling(doubSummaryBandwidth / 500));                                 //Calculated Total SPHD Number (4120 or 4160)

                int intHydrasOnFilter = 4;                                      //Default = 4160
                if (strFilterModel == "4120") intHydrasOnFilter = 3;            // 4120


                int intBalancersFromFiltersAndHydras = CalculateDevicesQuantity(intTotalFiltersFromBw * intHydrasOnFilter, 16);
                int intBalancersFromBw = CalculateDevicesQuantity(intTotalFiltersFromBw, intHydrasOnFilter);
                if (intTotalFiltersFromBw == 1) intBalancersFromBw = 0;


                //Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~");






                //~~~~~~~~~~~~~~~~~~~~~~~   Read Cells in Reference Excel File   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                //~~~~~~~~~~~~~~~   Variables for Excel File Content Read   ~~~~~~~~~~~~~~~~

                int intCurrentJournalItem = 0;
                int intLocalPortCounter = 0;
                int intCurrentLanPortsCounter;
                int intCurrentWanPortsCounter;
                int intCurrent100mBalancerPort;
                int intTotal100mBalancerPorts;

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

                bool boolLanObjectNotFound;
                bool boolWanObjectNotFound;

                int intCurrentGlobalDeviceIndex = 0;
                int intCurrentGlobalPortIndex = 0;
                int intCurrentGlobalDeviceIndexForPort = 0;
                //int intCurrentDeviceInGroup;
                int intGlobalCableCounter = 0;
                int intLinkCounter100 = 0;
                int intLinkCounter40 = 0;
                int intLinkCounter10 = 0;
                int intLinkCounter1Fiber = 0;
                int intLinkCounter1Copper = 0;
                int intLinkCounterEol = 0;

                int intCurrentOverallLinkNumber = 0;

                //Console.WriteLine($"Total Rows: {intTotalRows}");
                //Console.WriteLine("~~~~~~~~~~~  CHECK   ~~~~~~~~~~~~~~~~");

                //~~~~~~~~~~~~~~~~  Start For Cycle (pass through all rows in file) ~~~~~~~~~~~~~~~~~~~~~
                for (int intCurrentRow = 4; intCurrentRow <= intTotalRows; intCurrentRow++)
                {
                    //Console.WriteLine($"Current Row: {intCurrentRow}");
                    boolLanObjectNotFound = true;
                    boolWanObjectNotFound = true;
                    strCurrentLinkType = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 3]).Value2.ToString();
                    strCurrentLanHostname = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 4]).Value2.ToString();
                    strCurrentLanPortName = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 5]).Value2.ToString();
                    strCurrentWanHostname = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 8]).Value2.ToString();
                    strCurrentWanPortName = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 9]).Value2.ToString();
                    strLanOdfLocation = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 12]).Value2.ToString();
                    strLanOdfPort = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 13]).Value2.ToString();
                    strWanOdfLocation = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 14]).Value2.ToString();
                    strWanOdfPort = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 15]).Value2.ToString();
                    intCurrentOverallLinkNumber++;

                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~    Count Ports of Each Type    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
                        case "1Fiber":
                            intLinkCounter1Fiber++;
                            break;
                        case "1Copper":
                            intLinkCounter1Copper++;
                            break;
                        case "10_EOL":
                            intLinkCounterEol++;            // Потом удалить!
                            break;
                    }

                    //////////////////////////////////////////// LAN Devices ////////////////////////////////////////////////

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

                    if (boolLanObjectNotFound)
                    {
                        listLanDevices.Add(new Dictionary<string, object>());
                        intCurrentGlobalDeviceIndex++;
                        listLanDevices[listLanDevices.Count - 1].Add("Device_Name", (strCurrentLanHostname));
                        listLanDevices[listLanDevices.Count - 1].Add("Ports_Number", (1));
                        listLanDevices[listLanDevices.Count - 1].Add("Device_Index", intCurrentGlobalDeviceIndex);
                        intCurrentGlobalDeviceIndexForPort = intCurrentGlobalDeviceIndex;
                    };

                    ///////////////////////////////////// LAN Ports   ////////////////////////////////////////////////////

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
                        case "1Fiber":
                            intLocalPortCounter = 2 * intLinkCounter1Fiber - 1;
                            break;
                        case "1Copper":
                            intLocalPortCounter = 2 * intLinkCounter1Copper - 1;
                            break;
                        case "10_EOL":
                            intLocalPortCounter = 2 * intLinkCounterEol - 1;                                                //Убрать после того как переделаю логику
                            break;
                    }

                    //Console.WriteLine($"~~~~~~~~~~~~~~~~~~~~~~");
                    listLanPorts[listLanPorts.Count - 1].Add("Device_Index", (intCurrentGlobalDeviceIndexForPort));
                    listLanPorts[listLanPorts.Count - 1].Add("Port_Name", (strCurrentLanPortName));
                    listLanPorts[listLanPorts.Count - 1].Add("Port_Index", (intLocalPortCounter));                          //~~~~~~~~~~~~~~~~~~    Поменять!   ~~~~~~~~~~~~~~
                    listLanPorts[listLanPorts.Count - 1].Add("Link_Type", strCurrentLinkType);
                    listLanPorts[listLanPorts.Count - 1].Add("Overall_Link_Number", intCurrentOverallLinkNumber);

                    //В КЖ
                    intCurrentJournalItem++;
                    intGlobalCableCounter++;
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0] = new Dictionary<string, string>();
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Port_ID", Convert.ToString(intLocalPortCounter));
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "л");
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Cable_Name", "ODF --- " + strBypassModel);
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_A_Name", strLanOdfLocation);
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_A_Port", strLanOdfPort);
                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Comment", strCurrentLanHostname + "\n" + Convert.ToString(strCurrentLanPortName));


                    //////////////////////////////////////// WAN Devices ////////////////////////////////////////////////

                    //Заполняем список словарей по LAN-линкам.

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

                    if (boolWanObjectNotFound)
                    {
                        listWanDevices.Add(new Dictionary<string, object>());
                        intCurrentGlobalDeviceIndex++;
                        listWanDevices[listWanDevices.Count - 1].Add("Device_Name", (strCurrentWanHostname));
                        listWanDevices[listWanDevices.Count - 1].Add("Ports_Number", (1));
                        listWanDevices[listWanDevices.Count - 1].Add("Device_Index", intCurrentGlobalDeviceIndex);
                        intCurrentGlobalDeviceIndexForPort = intCurrentGlobalDeviceIndex;


                    };

                    /////////////////////////////// WAN Ports   ////////////////////////////////////////////////////

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
                        case "1Fiber":
                            intLocalPortCounter = 2 * intLinkCounter1Fiber;
                            break;
                        case "1Copper":
                            intLocalPortCounter = 2 * intLinkCounter1Copper;
                            break;
                        case "10_EOL":
                            intLocalPortCounter = 2 * intLinkCounterEol;
                            break;
                    }

                    listWanPorts[listWanPorts.Count - 1].Add("Device_Index", (intCurrentGlobalDeviceIndexForPort));
                    listWanPorts[listWanPorts.Count - 1].Add("Port_Name", (strCurrentWanPortName));
                    listWanPorts[listWanPorts.Count - 1].Add("Port_Index", (intLocalPortCounter));          //~~~~~~~~~~~~~~~~~~    Поменять!   ~~~~~~~~~~~~~~
                    listWanPorts[listWanPorts.Count - 1].Add("Link_Type", strCurrentLinkType);
                    listWanPorts[listWanPorts.Count - 1].Add("Overall_Link_Number", intCurrentOverallLinkNumber);

                    //В КЖ
                    intCurrentJournalItem++;
                    //intGlobalCableCounter++;
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1] = new Dictionary<string, string>();
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Port_ID", Convert.ToString(intLocalPortCounter));
                    //arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Number", Convert.ToString(intTotalRows - 3 + intGlobalCableCounter));
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "в");
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Cable_Name", "ODF --- " + strBypassModel);
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_A_Name", strWanOdfLocation);
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_A_Port", strWanOdfPort);
                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Comment", strCurrentWanHostname + "\n" + Convert.ToString(strCurrentWanPortName));


                    //////////////  End Excel Pass  ///////////////////////////////

                };

                


                //close excel
                xlWorkbook.Close();
                xlApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);




                /////////////////////  Create Output CJ Excel File  ///////////////////////////////

                //open excel
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

                //////////

                Excel.Range formatHeaders;

                formatHeaders = xlWorksheet31.get_Range("a2", "k3");
                formatHeaders.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.BurlyWood);

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                if (intLinkCounter100 > 0) Console.WriteLine($"Порты 100G: {intLinkCounter100}");
                if (intLinkCounter40 > 0) Console.WriteLine($"Порты 40G: {intLinkCounter40}");
                if (intLinkCounter10 > 0) Console.WriteLine($"Порты 10G: {intLinkCounter10}");
                if (intLinkCounter1Fiber > 0) Console.WriteLine($"Порты 1G (оптика): {intLinkCounter1Fiber}");
                if (intLinkCounter1Copper > 0) Console.WriteLine($"Порты 1G (медь): {intLinkCounter1Copper}");

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


                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Определение количества байпасов каждого типа    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  

                int intBypassPortDivider;                                                                    //EOL?               
                if (boolEolBypass) intBypassPortDivider = 4;
                else intBypassPortDivider = 6;

                int intTotalIs100Bypasses = CalculateDevicesQuantity(intLinkCounter100, 2);
                int intTotalIs40Bypasses = CalculateDevicesQuantity(intLinkCounter40, 3);
                int intTotalIs10Bypasses = CalculateDevicesQuantity(intLinkCounter10, intBypassPortDivider);
                int intTotalIs1FiberBypasses = CalculateDevicesQuantity(intLinkCounter1Fiber, 4);
                int intTotalIs1CopperBypasses = CalculateDevicesQuantity(intLinkCounter1Fiber, 4);
                int intTotalIbs1upBypasses = CalculateDevicesQuantity(intLinkCounter10, 4);

                /////////////////////////////   Потом удалить!  //////////////////////////////////////





                //~~~~~~~~~~~~~~~~~~~~~~~~~~~   Console output bypasses' quantity   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                if (intTotalIs100Bypasses > 0) Console.WriteLine($"IS100: {intTotalIs100Bypasses}");
                if (intTotalIs40Bypasses > 0 | (intTotalIs10Bypasses > 0) && !boolEolBypass) Console.WriteLine($"IS40 Number: {intTotalIs40Bypasses} + {intTotalIs10Bypasses} = {intTotalIs40Bypasses + intTotalIs10Bypasses}");
                if ((intTotalIs10Bypasses > 0) && boolEolBypass) Console.WriteLine($"IBS1U Number: {intTotalIs10Bypasses}");
                if (intTotalIs1FiberBypasses > 0 | intTotalIs1CopperBypasses > 0) Console.WriteLine($"IBS1U Number: {intTotalIs1FiberBypasses} + {intTotalIs1CopperBypasses} = {intTotalIs1FiberBypasses + intTotalIs1CopperBypasses}");


                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                double doubStartPointNextShapeX;
                double doubStartPointNextShapeY;


                double doubDeviceStartPointX = 0;
                double doubDeviceStartPointY;
                double doubDeviceEndPointX;
                double doubDeviceEndPointY;

                int intPortsOnDevice;

                ///////////////////////////////////////////////////////////////////////////////////


                ////////////////////////////////    Calculate Balancers Quantity   ///////////////////////////////////////




                bool boolNoBalancer = false;
                if (intLinkCounter10 > 0 & intLinkCounter10 <= 8) boolNoBalancer = true;

                if (boolNoBalancer) Console.WriteLine("Прямая коммутация байпасов к фильтру. Без балансировщиков.");

                int intBalancersFromPorts = CalculateDevicesQuantity(2 * intLinkCounter100 + 2 * intLinkCounter40 + intLinkCounter10 / 2, 16);
                int intBalancersFinalQuantity = Math.Max(intBalancersFromBw, intBalancersFromPorts);


                doubStartPointNextShapeX = 1;
                doubStartPointNextShapeY = 1;

                int intListCurrentIndex;

                intListCurrentIndex = 0;

                foreach (Dictionary<string, object> dictLanDevices in listLanDevices)
                {
                    intPortsOnDevice = Convert.ToInt32(dictLanDevices["Ports_Number"]);
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

                // Console.WriteLine($"~~~~~~~~~~~~~~~~~~~    WAN Device Dictionaries  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");

                doubStartPointNextShapeX = 1;
                doubStartPointNextShapeY = intTotalIs100Bypasses * 6 + intTotalIs40Bypasses * 8 + intTotalIs10Bypasses * 10 + intTotalIs1FiberBypasses * 6 + intTotalIs1CopperBypasses * 6; // + intTotalIbs1upBypasses * 6;


                intListCurrentIndex = 0;

                foreach (Dictionary<string, object> dictWanDevices in listWanDevices)
                {

                    intPortsOnDevice = Convert.ToInt32(dictWanDevices["Ports_Number"]);
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




                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////










                //~~~~~~~~~~~~~~~~~~~~~ Create Visio File   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Visio.Application app = new Visio.Application();
                Visio.Document doc = app.Documents.Add("");
                Visio.Page page1 = doc.Pages[1];
                page1.Name = "Схема Организации Связи";

                //~~~~~~~~~~~~~~~~~~~   Create Selection Variable   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Visio.Selection vsoSelection;
                Visio.Window vsoWindow;
                vsoWindow = app.ActiveWindow;




                /////////////////////////////////// Draw Rectangles //////////////////////////////////////////////


                double doubNextPortStartPointX;
                double doubNextPortStartPointY;

                double doubLanLastX;
                double doubWanLastX;

                string strShapeName;
                string strPortName;


                int intCurrentLanChassis = 0;
                int intCurrentWanChassis = 0;

                intCurrentOverallLinkNumber = 0;

                /////////////////////////////////// Draw LAN Chassis //////////////////////////////////////////////

                foreach (Dictionary<string, object> dictLanDevices in listLanDevices)
                {
                    intCurrentLanChassis++;
                    doubDeviceStartPointX = Convert.ToDouble(dictLanDevices["StartX"]);
                    doubDeviceStartPointY = Convert.ToDouble(dictLanDevices["StartY"]);
                    doubDeviceEndPointX = Convert.ToDouble(dictLanDevices["EndX"]);
                    doubDeviceEndPointY = Convert.ToDouble(dictLanDevices["EndY"]);

                    strShapeName = Convert.ToString(dictLanDevices["Device_Name"]);

                    listShapesLanDevices.Add(page1.DrawRectangle(doubDeviceStartPointX, doubDeviceStartPointY, doubDeviceEndPointX + 0.1, doubDeviceEndPointY));
                    listShapesLanDevices[listShapesLanDevices.Count - 1].Text = strShapeName;
                    listShapesLanDevices[listShapesLanDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(169,169,169)";

                    /////////////////////////////////// Draw LAN Ports //////////////////////////////////////////////

                    doubNextPortStartPointX = doubDeviceStartPointX - 0.1;
                    doubNextPortStartPointY = doubDeviceEndPointY + 0.2;

                    foreach (Dictionary<string, object> dictLanPorts in listLanPorts)
                    {
                        if (Convert.ToInt32(dictLanDevices["Device_Index"]) == Convert.ToInt32(dictLanPorts["Device_Index"]))
                        {
                            strPortName = Convert.ToString(dictLanPorts["Port_Name"]);
                            intCurrentOverallLinkNumber++;
                            arrShapesLanPorts[intCurrentOverallLinkNumber] = page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.6, doubNextPortStartPointY + 0.2);
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Data1 = Convert.ToString(dictLanPorts["Port_Index"]);
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Data2 = Convert.ToString(dictLanPorts["Link_Type"]);
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Data3 = Convert.ToString(intCurrentLanChassis);
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Text = strPortName;
                            arrShapesLanPorts[intCurrentOverallLinkNumber].Rotate90();
                            arrShapesLanPorts[intCurrentOverallLinkNumber].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            doubNextPortStartPointX += 0.2;
                        };
                    };
                    doubLanLastX = doubDeviceEndPointX;
                };

                intCurrentOverallLinkNumber = 0;

                /////////////////////////////////// Draw WAN Chassis //////////////////////////////////////////////

                double doubOldDevicesEndPointX;

                foreach (Dictionary<string, object> dictWanDevices in listWanDevices)
                {
                    intCurrentWanChassis++;
                    doubDeviceStartPointX = Convert.ToDouble(dictWanDevices["StartX"]);
                    doubDeviceStartPointY = Convert.ToDouble(dictWanDevices["StartY"]);
                    doubDeviceEndPointX = Convert.ToDouble(dictWanDevices["EndX"]);
                    doubDeviceEndPointY = Convert.ToDouble(dictWanDevices["EndY"]);

                    strShapeName = Convert.ToString(dictWanDevices["Device_Name"]);

                    listShapesWanDevices.Add(page1.DrawRectangle(doubDeviceStartPointX, doubDeviceStartPointY, doubDeviceEndPointX, doubDeviceEndPointY));
                    listShapesWanDevices[listShapesWanDevices.Count - 1].Text = strShapeName;
                    listShapesWanDevices[listShapesWanDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(169,169,169)";

                    doubOldDevicesEndPointX = doubDeviceEndPointX;

                    //listLanPorts[listLanPorts.Count - 1].Add("Overall_Link_Number", intCurrentOverallLinkNumber);

                    /////////////////////////////////// Draw WAN Ports //////////////////////////////////////////////

                    doubNextPortStartPointX = doubDeviceStartPointX - 0.1;
                    doubNextPortStartPointY = doubDeviceStartPointY - 0.2;

                    foreach (Dictionary<string, object> dictWanPorts in listWanPorts)
                    {
                        if (Convert.ToInt32(dictWanDevices["Device_Index"]) == Convert.ToInt32(dictWanPorts["Device_Index"]))
                        {
                            strPortName = Convert.ToString(dictWanPorts["Port_Name"]);
                            intCurrentOverallLinkNumber++;
                            arrShapesWanPorts[intCurrentOverallLinkNumber] = page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.6, doubNextPortStartPointY - 0.2);
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Data1 = Convert.ToString(dictWanPorts["Port_Index"]);
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Data2 = Convert.ToString(dictWanPorts["Link_Type"]);
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Data3 = Convert.ToString(intCurrentLanChassis);
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Text = strPortName;
                            arrShapesWanPorts[intCurrentOverallLinkNumber].Rotate90();
                            arrShapesWanPorts[intCurrentOverallLinkNumber].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            doubNextPortStartPointX += 0.2;
                        };
                    };
                    doubWanLastX = doubDeviceEndPointX;
                };

                //Фиксация точки Y, с которой рисовать следующие устройства.
                double intdoubTopLineY = 60;

                //Вычисление общего количества линков.
                int intTotalOverallLinkNumber = intCurrentOverallLinkNumber;

                // Выставляем указатель номера кабеля в КЖ на соедующий номер после перебора всех оплинков.
                //intGlobalCableCounter = (intTotalRows - 3) * 2;

                intGlobalCableCounter = 0;

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                List<Visio.Shape> listDeviceMgmtPorts = new List<Visio.Shape>();                                                                //Add MGMT Ports List (for all TSPU devices)





                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                int intCurrentBypassPort;

                double doubUpperStartPoint = doubStartPointNextShapeY - 1.5;

                //int intMaximumLinksOnAllBypasses;
                //int intMaximumLinksOnSingleBypass;
                int intDifference;


                ////////////////////////////////////////////////////// Draw Bypass IS100 Chassis ////////////////////////////////////////////////////////////////

                List<Visio.Shape> listShapesBypass100Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass100NetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass100MonPorts = new List<Visio.Shape>();

                Visio.Shape[,] arrShapesBypass100MonPorts = new Visio.Shape[10, 200];
                Visio.Shape[,] arrShapesBypass100NetPorts = new Visio.Shape[200, 10];

                Visio.Shape[,,] arrShapesBypass10_MonPorts = new Visio.Shape[200, 200, 10];
                Visio.Shape[,] arrShapesBypass10_NetPorts = new Visio.Shape[200, 10];

                Visio.Shape[,] arrShapesBypass40MonPorts = new Visio.Shape[10, 100];
                Visio.Shape[,] arrShapesFilterPorts = new Visio.Shape[200, 20];
                Visio.Shape[,] arrFilterHydraConnectors = new Visio.Shape[200, 10];
                Visio.Shape[,] arrBypassIs40HydraConnectors = new Visio.Shape[200, 200];



                int intTotalBalancers = Convert.ToInt32(strBalancerNumberFromInput);
                string strCurrentDeviceHostname;
                string strCurrentPortName;
                double doubBypassEndX;
                //doubStartPointNextShapeX = Math.Max(listShapesLanPorts.Count, listShapesWanPorts.Count) * 1.2;
                doubStartPointNextShapeX = intCurrentOverallLinkNumber * 1.2;
                doubStartPointNextShapeY -= 1.5;
                intCurrentBypassPort = 0;
                int intCurrentBalancerChassis = 0;
                int[] arrCurrentUplinkPortInBalancer = new int[20];


                // Обнуляем поинтеры для каждого балансировщика
                for (int intCurrentBalancer = 1; intCurrentBalancer <= intTotalBalancers; intCurrentBalancer++)
                {
                    arrCurrentUplinkPortInBalancer[intCurrentBalancer] = 16;
                }


                intCurrentOverallLinkNumber = 0;

                intCurrent100mBalancerPort = 0;
                //intTotal100mBalancerPorts;

                ///////////////////////////// Draw Bypass IS100 Chassis ////////////////////////////////////

                //Console.WriteLine($"На Байпасах:");

                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs100Bypasses; intCurrentBypassDevice++)
                {

                    listShapesBypass100Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 1.3));
                    strCurrentDeviceHostname = "IS100 (" + intCurrentBypassDevice + ")";
                    listShapesBypass100Devices[listShapesBypass100Devices.Count - 1].Text = strCurrentDeviceHostname;
                    listShapesBypass100Devices[listShapesBypass100Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(255,228,225)";

                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGNT ETH";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MGNT ETH";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(intCurrentBypassDevice);
                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    doubStartPointNextShapeY -= 2;

                    ///////////////////////////// Draw Bypass IS100 Ports ////////////////////////////////////

                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Net-порты (Operator-Bypass) ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 2; intCurrentPortCounterInBypass++)
                    {
                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        intCurrentOverallLinkNumber++;
                        intCurrentBypassPort++;


                        //Отрисовка портов Net0
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/0";
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].Text = strCurrentPortName;
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrent100mBalancerPort++;

                        //Дозаполнение ячейки массива словарей LAN-Bypass
                        arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Name", strCurrentDeviceHostname);
                        arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Port", strCurrentPortName);
                        list_CableJournal_LAN_Bypass.Add(new Dictionary<string, string>());

                        //Перенос словарей из нумерованного списка в ненумерованный - чтобы отработал foreach
                        foreach (string key in arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Keys)
                        {
                            list_CableJournal_LAN_Bypass[list_CableJournal_LAN_Bypass.Count - 1].Add(key, arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0][key]);
                        };




                        //Отрисовка портов Net1
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/1";
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].Text = strCurrentPortName;
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrent100mBalancerPort++;

                        //Дозаполнение ячейки массива словарей WAN-Bypass
                        arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Name", strCurrentDeviceHostname);
                        arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Port", strCurrentPortName);
                        list_CableJournal_WAN_Bypass.Add(new Dictionary<string, string>());
                        foreach (string key in arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Keys)
                        {
                            list_CableJournal_WAN_Bypass[list_CableJournal_WAN_Bypass.Count - 1].Add(key, arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1][key]);
                        };


                        //Соеденение двух Net-портов с LAN и WAN
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0].AutoConnect(arrShapesLanPorts[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1].AutoConnect(arrShapesWanPorts[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                        //Visio.Shape Connect0 = page1.DropConnected(arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 0], arrShapesLanPorts[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        //Visio.Shape Connect1 = page1.DropConnected(arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1], arrShapesWanPorts[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                        //Connect1.

                        //Раундробин (формула).
                        intCurrentBalancerChassis++;

                        //WAN-линки - в нечётные порты балансировщика (1, 3, 5, 7,...)
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/1";
                        //Console.WriteLine($"{strCurrentDeviceHostname}, {strCurrentPortName} --- Balancer {intCurrentBalancerChassis}, Port {arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 1}");
                        intCurrentPortInChassis = arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 1;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Text = strCurrentPortName;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        //КЖ линков WAN
                        intGlobalCableCounter++;
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0] = new Dictionary<string, string>();
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бб");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Device_A_Name", strCurrentDeviceHostname);

                        //LAN-линки - в чётные порты балансировщика (2, 4, 6, 8,...)
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                        //Console.WriteLine($"{strCurrentDeviceHostname}, {strCurrentPortName} --- Balancer {intCurrentBalancerChassis}, Port {arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 1}");
                        intCurrentPortInChassis = arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 2;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Text = strCurrentPortName;
                        arrShapesBypass100MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        //КЖ линков LAN
                        intGlobalCableCounter++;
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0] = new Dictionary<string, string>();
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бб");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Device_A_Name", strCurrentDeviceHostname);






                        //После отрисовки пары WAN-LAN сдвиг указателя на 2 порта.
                        arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] += 2;


                        //Console.WriteLine($"Балансировщик: {intCurrentBalancerChassis}, Порт: {intCurrentPortInChassis}");


                        //После полного прогона балансировщиков указатель перемещается в начало - к первому балансировщику.
                        if (intCurrentBalancerChassis == intTotalBalancers) intCurrentBalancerChassis = 0;

                        //Сдвиг вниз следующей фигуры байпаса
                        doubNextPortStartPointY -= 0.7;

                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    };


                };



                //Visio.Shape connector = page1.Drop()

                //currentStencil.Masters["Dynamic connector"], 4.50, 4.50);

                //Visio.Window stencilWindow = page1.Document.OpenStencilWindow();
                //Visio.Page currentPage = Microsoft.Office.Interop.VisOcx.DrawingControl.Equals.;
                //Visio.Document currentStencil = axDrawingControl1.Document.Application.Documents.OpenEx("Basic_U.vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);
                //int countStencils = currentStencil.Masters.Count;

                //Visio.Shape shape1 = page1.D

                //Set vsoSquareShape = ActiveWindow.Page.Drop(Documents("BASIC_U.VSS").Masters.ItemU("Square"), 4, 9)

                //vsoSquareShape = page1.Document.Masters.Application.AutoLayout;



                intTotal100mBalancerPorts = intCurrent100mBalancerPort;
                //Zenit
                ////////////////////////////////////////////////////// Draw Bypass IS40 Chassis (for 10G) ////////////////////////////////////////////////////////////////



                List<Visio.Shape> listShapesBypass10Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass10NetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass10MonPorts = new List<Visio.Shape>();

                List<Visio.Shape> listHydraLines = new List<Visio.Shape>();
                List<Visio.Shape> listBypassHydraConnectors = new List<Visio.Shape>();



                List<Visio.Shape> listShapesIbs1UpDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesIbs1upNetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesIbs1upMonPorts = new List<Visio.Shape>();



                int iCurrentOverallMonPort = 0;
                int intFilterPortNoBalancer = 0;
                int intBypassIs40CurrentHydra = 0;
                int intBypassIs40HydrasTotal = 0;
                bool boolLastBypassChassis = false;
                int intMgmtPortOnChassis;
                int intHydraEnd;

                intCurrentBalancerChassis = 0;
                intCurrentPortInChassis = 16;
                intCurrentOverallLinkNumber = 0;

                //10G для IS40

                // Сдвиг ряда фигур IS40 вниз
                doubStartPointNextShapeY -= 1;

                //Обнуление счётчика портов байпаса
                intCurrentBypassPort = 0;

                // Подсчёт количества занятых байпас-сегментов в IBS1UP. Чтобы завявить меньше MGMT-портов на шасси.
                intDifference = intTotalOverallLinkNumber % 4;


                Console.WriteLine($"Отрисовка байпасов начата.");

                //Проход по всем десяточным байпасам (IS40 или IBS1UP)
                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs10Bypasses; intCurrentBypassDevice++)
                {

                    if (!boolEolBypass)         //  Байпас IS40
                    {
                        //////  IS40    //////
                        listShapesBypass10Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 4.1));
                        strCurrentDeviceHostname = "IS40 (" + intCurrentBypassDevice + ")";
                        listShapesBypass10Devices[listShapesBypass10Devices.Count - 1].Text = strCurrentDeviceHostname;
                        listShapesBypass10Devices[listShapesBypass10Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(219,112,147)";
                        //listShapesBypass10Devices[listShapesBypass10Devices.Count - 1].Data3 = Convert.ToString(intCurrentBypassDevice);

                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGNT ETH";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MGNT ETH";
                        //listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = Convert.ToString(intCurrentBypassDevice);

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        doubStartPointNextShapeY -= 5.4;



                        //Придумать, как заполнять и 40G и 10G и 1G под одним шасси



                        /////////////////////////// Draw Bypass IS40 Ports (for 10G) ////////////////////////////////////

                        intCurrentJournalItem = 0;

                        for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 3; intCurrentPortCounterInBypass++)
                        {
                            for (int intCurrentSubslotCounterInBypass = 1; intCurrentSubslotCounterInBypass <= 2; intCurrentSubslotCounterInBypass++)
                            {
                                intCurrentOverallLinkNumber++;
                                intCurrentBypassPort++;

                                //Net X/X/0
                                strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/0";
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Text = strCurrentPortName;
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                                //Соеденение двух Net-портов с LAN и WAN. Не соединяем Net-порты, к которым не приходят линки.
                                if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                                {
                                    //Соединительная линия.
                                    arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].AutoConnect(arrShapesLanPorts[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                                    //Дозаполнение ячейки массива словарей LAN-Bypass
                                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Name", strCurrentDeviceHostname);
                                    arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Port", strCurrentPortName);
                                    list_CableJournal_LAN_Bypass.Add(new Dictionary<string, string>());
                                    foreach (string key in arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Keys)
                                    {
                                        list_CableJournal_LAN_Bypass[list_CableJournal_LAN_Bypass.Count - 1].Add(key, arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0][key]);
                                    };
                                };

                                //Net X/X/1
                                strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/1";
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Text = strCurrentPortName;
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                                //Соеденение двух Net-портов с LAN и WAN. Не соединяем Net-порты, к которым не приходят линки.
                                if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                                {
                                    //Соединительная линия.
                                    arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].AutoConnect(arrShapesWanPorts[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                                    //Дозаполнение ячейки массива словарей WAN-Bypass
                                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Name", strCurrentDeviceHostname);
                                    arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Port", strCurrentPortName);
                                    list_CableJournal_WAN_Bypass.Add(new Dictionary<string, string>());
                                    foreach (string key in arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Keys)
                                    {
                                        list_CableJournal_WAN_Bypass[list_CableJournal_WAN_Bypass.Count - 1].Add(key, arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1][key]);
                                    };
                                };




                                //Mon-порты
                                //Если включение IS40 сразу к фильтру, intCurrentBalancerChassis = 0, intCurrentPortInChassis номер порта фильтра)

                                if ((intCurrentBypassPort - 1) % 2 == 0 && !boolNoBalancer) intCurrentPortInChassis++;
                                if ((intCurrentBypassPort - 1) % 32 == 0 && !boolNoBalancer)
                                {
                                    intCurrentBalancerChassis++;
                                    intCurrentPortInChassis = 17;
                                };

                                //intCurrentBypassPort++;





                                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                //Mon/X/1
                                iCurrentOverallMonPort++;
                                strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/1";

                                if (!boolNoBalancer && iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
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

                                    //Console.WriteLine($"{strCurrentDeviceHostname}, {strCurrentPortName} ----- Балансер: {intCurrentBalancerChassis}, Порт: {intCurrentPortInChassis}");
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass].Data1 = Convert.ToString(intCurrentBypassDevice);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass].Data2 = Convert.ToString(intCurrentBypassPort);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass].Text = strCurrentPortName;
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                    intGlobalCableCounter++;

                                    /*
                                    if (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                                    {

                                    };
                                    */

                                    //Console.WriteLine($"Балансировщик: {intCurrentBalancerChassis}, Порт: {intCurrentPortInChassis}, Конец гидры: {intHydraEnd}, Строка КЖ: {intGlobalCableCounter}");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd] = new Dictionary<string, string>();
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бб");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + strCableInHydra);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Device_A_Name", strCurrentDeviceHostname);
                                }
                                else if (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                                {
                                    intFilterPortNoBalancer++;
                                    strCableInHydra = "";
                                    //Console.WriteLine($"Порт фильтра: {intFilterPortNoBalancer}");
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0].Text = strCurrentPortName;
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                    //Console.WriteLine($"Номер порта фильтра: {intFilterPortNoBalancer * 2 - 1}");
                                    intGlobalCableCounter++;
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0] = new Dictionary<string, string>();
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бб");
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                                };








                                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                // //Mon/X/0
                                iCurrentOverallMonPort++;
                                strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/0";

                                if (!boolNoBalancer && iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
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

                                    //Console.WriteLine($"{strCurrentDeviceHostname}, {strCurrentPortName} ----- Балансер: {intCurrentBalancerChassis}, Порт: {intCurrentPortInChassis}");
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass * 2] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass * 2].Data1 = Convert.ToString(intCurrentBypassDevice);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass * 2].Data2 = Convert.ToString(intCurrentBypassPort);
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass * 2].Text = strCurrentPortName;
                                    arrShapesBypass10_MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis, intCurrentSubslotCounterInBypass * 2].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                                    intGlobalCableCounter++;
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd] = new Dictionary<string, string>();
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бб");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + strCableInHydra);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Device_A_Name", strCurrentDeviceHostname);


                                }
                                else if (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                                {
                                    //intFilterPortNoBalancer++;
                                    strCableInHydra = "";
                                    // Console.WriteLine($"Порт фильтра: {intFilterPortNoBalancer}");
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0].Text = strCurrentPortName;
                                    arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                                    intGlobalCableCounter++;
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0] = new Dictionary<string, string>();
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бф");
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                                    arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Device_A_Name", strCurrentDeviceHostname);

                                };



                                /*
                                if (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                                {
                                    intGlobalCableCounter++;
                                    Console.WriteLine($"Балансировщик: {intCurrentBalancerChassis}, Порт: {intCurrentPortInChassis}, Конец гидры: {intHydraEnd}, Строка КЖ: {intGlobalCableCounter}");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd] = new Dictionary<string, string>();
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Number", Convert.ToString(intGlobalCableCounter));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + strCableInHydra);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Device_A_Name", strCurrentDeviceHostname);
                                };
                                */
                                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                //КЖ Байпас-Фильтр
                                //if (boolNoBalancer || iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2)
                                //Избавиться от списка в пользу массива. Список появится при заполнении портов фильтра.

                                /*
                                if (boolNoBalancer)
                                {
                                    list_CableJournal_Bypass_Filter.Add(new Dictionary<string, string>());
                                    intCurrentJournalItem++;
                                    //intGlobalCableCounter++;
                                    //Console.WriteLine($"Row: {intCurrentJournalItem}, Device: {strCurrentDeviceHostname}, Port: {strCurrentPortName}");
                                    list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter));
                                    list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                                    list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                    list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_A_Name", strCurrentPortName + strCableInHydra);
                                    list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Device_A_Name", strCurrentDeviceHostname);
                                };
                                */

                                ///////////////     Draw Hydra Connector    //////////////////////////
                                if (!boolNoBalancer && (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2))
                                {
                                    listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.4));
                                    listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.2));

                                    if (listHydraLines.Count == 4)
                                    {
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
                                        //Console.WriteLine($"Номер Байпаса: {intCurrentBypassDevice}");
                                        listHydraLines.Clear();
                                    };
                                };
                                doubBypassEndX = doubStartPointNextShapeX + 1;
                                doubNextPortStartPointY -= 0.7;                 // Move down to next port
                            };

                        };

                    }
                    else
                    {
                        //////////////////////////////////////////////////// Draw Bypass IBS1UP Chassis (10G) ////////////////////////////////////////////////////////////////
                        ///
                        listShapesIbs1UpDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 2.7));
                        strCurrentDeviceHostname = "IBS1UP (" + intCurrentBypassDevice + ")";
                        listShapesIbs1UpDevices[listShapesIbs1UpDevices.Count - 1].Text = strCurrentDeviceHostname;
                        listShapesIbs1UpDevices[listShapesIbs1UpDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(135,206,250)";
                        listShapesIbs1UpDevices[listShapesIbs1UpDevices.Count - 1].Data3 = strCurrentDeviceHostname;

                        double doubMgmtPortX = doubStartPointNextShapeX + 0.3;
                        double doubMgmtPortY = doubStartPointNextShapeY + 0.12;

                        if (intCurrentBypassDevice == intTotalIs10Bypasses) boolLastBypassChassis = true;

                        if (boolLastBypassChassis) intMgmtPortOnChassis = 4 - intDifference;
                        else intMgmtPortOnChassis = 4;


                        for (int intIbs1upCurrentPortMgmt = 1; intIbs1upCurrentPortMgmt <= intMgmtPortOnChassis; intIbs1upCurrentPortMgmt++)
                        {
                            listDeviceMgmtPorts.Add(page1.DrawRectangle(doubMgmtPortX, doubMgmtPortY, doubMgmtPortX + 0.45, doubMgmtPortY + 0.2));
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGMT-" + intIbs1upCurrentPortMgmt;
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MANAGEMENT-" + intIbs1upCurrentPortMgmt;
                            listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = strCurrentDeviceHostname;

                            doubMgmtPortX += 0.2;
                        };

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;
                        doubStartPointNextShapeY -= 4;




                        /////////////////////////// Draw IBS1UP Ports (10G) ////////////////////////////////////

                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~    Operator-Bypass ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                        for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 4; intCurrentPortCounterInBypass++)
                        {
                            //Net-порты
                            intCurrentOverallLinkNumber++;
                            intCurrentBypassPort++;

                            //NetX/X/0
                            strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/0";
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].Text = strCurrentPortName;
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                            if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                            {
                                //Соединительная линия.
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 0].AutoConnect(arrShapesLanPorts[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                                //Дозаполнение ячейки массива словарей LAN-Bypass
                                arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Name", strCurrentDeviceHostname);
                                arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Add("Side_B_Port", strCurrentPortName);
                                list_CableJournal_LAN_Bypass.Add(new Dictionary<string, string>());
                                foreach (string key in arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0].Keys)
                                {
                                    list_CableJournal_LAN_Bypass[list_CableJournal_LAN_Bypass.Count - 1].Add(key, arrCableJournal_LAN_Bypass[intCurrentOverallLinkNumber, 0][key]);
                                };
                            };


                            //NetX/X/1
                            strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/1";
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Data2 = Convert.ToString(intCurrentBypassPort);
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].Text = strCurrentPortName;
                            arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                            if (intCurrentOverallLinkNumber <= intTotalOverallLinkNumber)
                            {
                                //Соединительная линия.
                                arrShapesBypass10_NetPorts[intCurrentOverallLinkNumber, 1].AutoConnect(arrShapesWanPorts[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                                //Дозаполнение ячейки массива словарей WAN-Bypass
                                arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Name", strCurrentDeviceHostname);
                                arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Add("Side_B_Port", strCurrentPortName);
                                list_CableJournal_WAN_Bypass.Add(new Dictionary<string, string>());
                                foreach (string key in arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1].Keys)
                                {
                                    list_CableJournal_WAN_Bypass[list_CableJournal_WAN_Bypass.Count - 1].Add(key, arrCableJournal_WAN_Bypass[intCurrentOverallLinkNumber, 1][key]);
                                };
                            };

                            //Mon-порты

                            if ((intCurrentBypassPort - 1) % 2 == 0 && !boolNoBalancer) intCurrentPortInChassis++;
                            if ((intCurrentBypassPort - 1) % 32 == 0 && !boolNoBalancer)
                            {
                                intCurrentBalancerChassis++;
                                intCurrentPortInChassis = 17;
                            };

                            //------------------------------------------------------------------------------------------------------------------

                            iCurrentOverallMonPort++;
                            strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                            //Console.WriteLine($"!!!!!! Checked 1 !!!!!!");
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

                                //Console.WriteLine($"Номер Порта Фильтра: {intCurrentPortInChassis}, Имя Порта IBS1UP: {strCurrentPortName}, {strCableInHydra}");
                                intGlobalCableCounter++;
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd] = new Dictionary<string, string>();
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Row", Convert.ToString(intCurrentJournalItem));
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бб");
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + strCableInHydra);
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Device_A_Name", strCurrentDeviceHostname);

                            }
                            else
                            {
                                intFilterPortNoBalancer++;
                                strCableInHydra = "";
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0].Text = strCurrentPortName;
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                                intGlobalCableCounter++;
                                //Console.WriteLine($"Номер Порта Фильтра: {intFilterPortNoBalancer * 2}, Имя Порта IBS1UP: {strCurrentPortName}, {strCableInHydra}");
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0] = new Dictionary<string, string>();
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бф");
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Cable_Name", strBypassModel + "---" + strFilterModel);
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                            };






                            iCurrentOverallMonPort++;
                            strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/1";

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

                                //Console.WriteLine($"Балансировщик: {intCurrentBalancerChassis}, Номер Порта: {intCurrentPortInChassis}, Имя Порта IBS1UP: {strCurrentPortName}, {strCableInHydra}");
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd] = new Dictionary<string, string>();
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Row", Convert.ToString(intCurrentJournalItem));
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бб");
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + strCableInHydra);
                                arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, intHydraEnd].Add("Device_A_Name", strCurrentDeviceHostname);
                            }
                            else
                            {
                                strCableInHydra = "";
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1);
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0].Data1 = Convert.ToString(intCurrentBypassDevice);
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0].Data2 = Convert.ToString(intCurrentBypassPort);
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0].Text = strCurrentPortName;
                                arrShapesBypass10_MonPorts[0, intFilterPortNoBalancer * 2 - 1, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                                intGlobalCableCounter++;
                                //Console.WriteLine($"Номер Порта Фильтра: {intFilterPortNoBalancer * 2 - 1}, Имя Порта IBS1UP: {strCurrentPortName}, {strCableInHydra}");
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0] = new Dictionary<string, string>();
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бб");
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Cable_Name", strBypassModel + "---" + strFilterModel);
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                                arr_CableJournal_Bypass_Balancer[0, intFilterPortNoBalancer * 2 - 1, 0].Add("Device_A_Name", strCurrentDeviceHostname);

                            };


                            ///////////////     Draw Hydra Connector    //////////////////////////
                            if (!boolNoBalancer && (iCurrentOverallMonPort <= intTotalOverallLinkNumber * 2))
                            {
                                listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.4));
                                listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 2.5 + 0.1, doubNextPortStartPointY - 0.2));

                                if (listHydraLines.Count == 4)
                                {
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
                                };

                            };



                            doubNextPortStartPointY -= 0.7;
                        };




                    };
                };




                intBypassIs40HydrasTotal = intBypassIs40CurrentHydra;

                //Console.WriteLine($"Количество гидр с IS40: {intBypassIs40HydrasTotal}");

























                ////////////////////////////////////////////////////// Draw Bypass IS40 Chassis (for 40G) ////////////////////////////////////////////////////////////////

                List<Visio.Shape> listShapesBypass40Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass40NetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass40MonPorts = new List<Visio.Shape>();

                //doubStartPointNextShapeX = doubStartPointNextShapeX;
                //doubStartPointNextShapeY -= 1;










                intCurrentBypassPort = 0;

                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs40Bypasses; intCurrentBypassDevice++)
                {
                    listShapesBypass40Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 2));
                    strCurrentDeviceHostname = "IS40-" + intCurrentBypassDevice;
                    listShapesBypass40Devices[listShapesBypass40Devices.Count - 1].Text = strCurrentDeviceHostname;
                    listShapesBypass40Devices[listShapesBypass40Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(255,165,0)";

                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGMT";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();

                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    doubStartPointNextShapeY -= 2.7;

                    //Liver


                    /////////////////////////// Draw Bypass IS40 Ports (40G) ////////////////////////////////////


                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 3; intCurrentPortCounterInBypass++)
                    {
                        //Отрисовка портов Net0
                        intCurrentBypassPort++;
                        listShapesBypass40NetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3));
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Text = strCurrentPortName;
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        //Заполнение словаря КЖ (LAN-IS40)
                        foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_LAN_Bypass)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Side_B_Name", strCurrentDeviceHostname);
                                dictCableRecord.Add("Side_B_Port", strCurrentPortName);
                                break;
                            };
                        };

                        /*
                        foreach (Dictionary<string, string> dictCableRecord in listCableJournal_4)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Bypass_Chassis", (strCurrentDeviceHostname));
                                dictCableRecord.Add("Bypass_LAN_Port", ("Net " + intCurrentPortCounterInBypass + "/0"));
                                dictCableRecord.Add("Bypass_WAN_Port", ("Net " + intCurrentPortCounterInBypass + "/1"));
                                break;
                            };
                        };
                        */

                        //Раундробин (формула).
                        intCurrentBalancerChassis++;
                        //Console.WriteLine($"{strCurrentDeviceHostname}, {strCurrentPortName} --- Balancer {intCurrentBalancerChassis}, Port {arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 2}");
                        //      !!! Отдебажить! Порты на балансерах заполнять либо с 0 либо после заполнения портов 100G !!!

                        /*
                        listShapesBypass40MonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3));
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Text = strCurrentPortName;
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        */

                        // LAN - линки - в чётные порты балансировщика(2, 4, 6, 8, ...)
                        intCurrentPortInChassis = arrCurrentUplinkPortInBalancer[intCurrentBalancerChassis] + 2;
                        arrShapesBypass40MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis] = page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3);
                        arrShapesBypass40MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data1 = Convert.ToString(intCurrentBypassDevice);
                        arrShapesBypass40MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Data2 = Convert.ToString(intCurrentBypassPort);
                        arrShapesBypass40MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].Text = strCurrentPortName;
                        arrShapesBypass40MonPorts[intCurrentBalancerChassis, intCurrentPortInChassis].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";





                        //Возможно, под удаление. Старый КЖ.
                        intGlobalCableCounter++;
                        list_CableJournal_Bypass_Filter.Add(new Dictionary<string, string>());
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Row", Convert.ToString(intCurrentJournalItem));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бф");
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Device_A_Name", strCurrentDeviceHostname);

                        //Новый КЖ - на стык с балансировщиками.
                        //Console.WriteLine($"Балансировщик {intCurrentBalancerChassis}, Порт {intCurrentPortInChassis}.");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0] = new Dictionary<string, string>();
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Row", Convert.ToString(intCurrentJournalItem));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бф");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        arr_CableJournal_Bypass_Balancer[intCurrentBalancerChassis, intCurrentPortInChassis, 0].Add("Device_A_Name", strCurrentDeviceHostname);
                        //Console.WriteLine($"Добавлено в словарь: [{intCurrentBalancerChassis}, {intCurrentPortInChassis}]");

                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        //Отрисовка портов Net1
                        intCurrentBypassPort++;
                        listShapesBypass40NetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1));
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/1";
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Text = strCurrentPortName;
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        /*
                        //Заполнение словаря КЖ (WAN-IS40)
                        foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_WAN_Bypass)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Side_B_Name", strCurrentDeviceHostname);
                                dictCableRecord.Add("Side_B_Port", strCurrentPortName);
                                break;
                            };
                        };
                        */





                        /*
                        foreach (Dictionary<string, string> dictCableRecord in listCableJournal_1)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Device_B_Name", (strCurrentDeviceHostname));
                                dictCableRecord.Add("Port_B_Name", (strCurrentPortName));
                                break;
                            };
                        };
                        */


                        /*
                        listShapesBypass40MonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.1));
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/1";
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Text = strCurrentPortName;
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        */

                        //Возможно, под удаление. Старый КЖ.
                        /*
                        intGlobalCableCounter++;
                        list_CableJournal_Bypass_Filter.Add(new Dictionary<string, string>());
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Row", Convert.ToString(intCurrentJournalItem));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Cable_Name", strBypassModel + " --- ELB-0133");
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Device_A_Name", strCurrentDeviceHostname);
                        */
                        //После полного прогона балансировщиков указатель перемещается в начало - к первому балансировщику.
                        if (intCurrentBalancerChassis == intTotalBalancers) intCurrentBalancerChassis = 0;                      //Отдебажить!

                        //Сдвиг вниз следующей фигуры байпаса
                        doubNextPortStartPointY -= 0.7;
                    };

                }



















                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                ////////////////////////////////////////////////////// Draw Bypass IBS1U Chassis (1G Fiber Ports) ////////////////////////////////////////////////////////////////

                List<Visio.Shape> listShapesBypass1FiberDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass1FiberNetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass1FiberMonPorts = new List<Visio.Shape>();

                //doubStartPointNextShapeX = doubDeviceStartPointX + 4;
                doubStartPointNextShapeY -= 1.5;

                intCurrentBypassPort = 0;

                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs1FiberBypasses; intCurrentBypassDevice++)
                {
                    listShapesBypass1FiberDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 2.7));
                    strCurrentDeviceHostname = "IBS1U-" + intCurrentBypassDevice;
                    listShapesBypass1FiberDevices[listShapesBypass1FiberDevices.Count - 1].Text = strCurrentDeviceHostname;
                    listShapesBypass1FiberDevices[listShapesBypass1FiberDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(255,228,225)";

                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGMT";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MANAGEMENT";

                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    doubStartPointNextShapeY -= 4;


                    /////////////////////////// Draw Bypass IBS1U Ports (Fiber) ////////////////////////////////////


                    //doubBypassEndX

                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 4; intCurrentPortCounterInBypass++)
                    {
                        intCurrentBypassPort++;
                        listShapesBypass1FiberNetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3));
                        listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Text = strCurrentPortName;
                        //listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0 (" + listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data2 + ")";
                        listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_LAN_Bypass)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Side_B_Name", strCurrentDeviceHostname);
                                dictCableRecord.Add("Side_B_Port", strCurrentPortName);
                                break;
                            };
                        };

                        /*
                        foreach (Dictionary<string, string> dictCableRecord in listCableJournal_1)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Device_B_Name", (strCurrentDeviceHostname));
                                dictCableRecord.Add("Port_B_Name", (strCurrentPortName));
                                break;
                            };
                        };
                        */

                        listShapesBypass1FiberMonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3));
                        listShapesBypass1FiberMonPorts[listShapesBypass1FiberMonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1FiberMonPorts[listShapesBypass1FiberMonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass1FiberMonPorts[listShapesBypass1FiberMonPorts.Count - 1].Text = strCurrentPortName;
                        listShapesBypass1FiberMonPorts[listShapesBypass1FiberMonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrentJournalItem++;
                        list_CableJournal_Bypass_Filter.Add(new Dictionary<string, string>());
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Row", Convert.ToString(intCurrentJournalItem));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Device_A_Name", strCurrentDeviceHostname);



                        intCurrentBypassPort++;
                        listShapesBypass1FiberNetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1));
                        listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/1";
                        listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Text = strCurrentPortName;
                        //listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0 (" + listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data2 + ")";
                        listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";


                        foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_WAN_Bypass)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Side_B_Name", strCurrentDeviceHostname);
                                dictCableRecord.Add("Side_B_Port", strCurrentPortName);
                                break;
                            };
                        };

                        /*
                        foreach (Dictionary<string, string> dictCableRecord in listCableJournal_1)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Device_B_Name", (strCurrentDeviceHostname));
                                dictCableRecord.Add("Port_B_Name", (strCurrentPortName));
                                break;
                            };
                        };
                        */

                        listShapesBypass1FiberMonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3));
                        listShapesBypass1FiberMonPorts[listShapesBypass1FiberMonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1FiberMonPorts[listShapesBypass1FiberMonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass1FiberMonPorts[listShapesBypass1FiberMonPorts.Count - 1].Text = strCurrentPortName;
                        listShapesBypass1FiberMonPorts[listShapesBypass1FiberMonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrentJournalItem++;
                        list_CableJournal_Bypass_Filter.Add(new Dictionary<string, string>());
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Row", Convert.ToString(intCurrentJournalItem));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Device_A_Name", strCurrentDeviceHostname);


                        doubNextPortStartPointY -= 0.7;
                    };






                }


                doubBypassEndX = doubStartPointNextShapeX + 1;
















                ////////////////////////////////////////////////////// Draw Bypass IBS1U Chassis (1G Copper Ports) ////////////////////////////////////////////////////////////////

                List<Visio.Shape> listShapesBypass1CopperDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass1CopperNetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass1CopperMonPorts = new List<Visio.Shape>();

                //doubStartPointNextShapeX = doubDeviceStartPointX + 4;
                doubStartPointNextShapeY -= 1.5;

                intCurrentBypassPort = 0;

                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs1CopperBypasses; intCurrentBypassDevice++)
                {
                    listShapesBypass1CopperDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 1, doubStartPointNextShapeY - 2.7));
                    strCurrentDeviceHostname = "IBS1U-" + intCurrentBypassDevice;
                    listShapesBypass1CopperDevices[listShapesBypass1CopperDevices.Count - 1].Text = strCurrentDeviceHostname;
                    listShapesBypass1CopperDevices[listShapesBypass1CopperDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(135,206,250)";

                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGMT";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MANAGEMENT";

                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    doubStartPointNextShapeY -= 4;





                    /////////////////////////// Draw Bypass IBS1U Ports (Copper) ////////////////////////////////////

                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~    Operator-Bypass ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 4; intCurrentPortCounterInBypass++)
                    {
                        intCurrentBypassPort++;
                        listShapesBypass1CopperNetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3));
                        listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Text = strCurrentPortName;
                        listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_LAN_Bypass)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Side_B_Name", strCurrentDeviceHostname);
                                dictCableRecord.Add("Side_B_Port", strCurrentPortName);
                                break;
                            };
                        };

                        /*
                        foreach (Dictionary<string, string> dictCableRecord in listCableJournal_1)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Device_B_Name", (strCurrentDeviceHostname));
                                dictCableRecord.Add("Port_B_Name", (strCurrentPortName));
                                break;
                            };
                        };
                        */

                        listShapesBypass1CopperMonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3));
                        listShapesBypass1CopperMonPorts[listShapesBypass1CopperMonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1CopperMonPorts[listShapesBypass1CopperMonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass1CopperMonPorts[listShapesBypass1CopperMonPorts.Count - 1].Text = strCurrentPortName;
                        listShapesBypass1CopperMonPorts[listShapesBypass1CopperMonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrentJournalItem++;
                        list_CableJournal_Bypass_Filter.Add(new Dictionary<string, string>());
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Row", Convert.ToString(intCurrentJournalItem));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Device_A_Name", strCurrentDeviceHostname);


                        intCurrentBypassPort++;
                        listShapesBypass1CopperNetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1));
                        listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Net " + intCurrentPortCounterInBypass + "/1";
                        listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Text = strCurrentPortName;
                        //listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0 (" + listShapesBypass1FiberNetPorts[listShapesBypass1FiberNetPorts.Count - 1].Data2 + ")";
                        listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";


                        foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_WAN_Bypass)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Side_B_Name", strCurrentDeviceHostname);
                                dictCableRecord.Add("Side_B_Port", strCurrentPortName);
                                break;
                            };
                        };

                        /*
                        foreach (Dictionary<string, string> dictCableRecord in listCableJournal_1)
                        {
                            if (dictCableRecord["Port_ID"] == listShapesBypass1CopperNetPorts[listShapesBypass1CopperNetPorts.Count - 1].Data2)
                            {
                                dictCableRecord.Add("Device_B_Name", (strCurrentDeviceHostname));
                                dictCableRecord.Add("Port_B_Name", (strCurrentPortName));
                                break;
                            };
                        };
                        */

                        listShapesBypass1CopperMonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 2, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 2.5, doubNextPortStartPointY - 0.3));
                        listShapesBypass1CopperMonPorts[listShapesBypass1CopperMonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1CopperMonPorts[listShapesBypass1CopperMonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        strCurrentPortName = "Mon " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass1CopperMonPorts[listShapesBypass1CopperMonPorts.Count - 1].Text = strCurrentPortName;
                        listShapesBypass1CopperMonPorts[listShapesBypass1CopperMonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrentJournalItem++;
                        list_CableJournal_Bypass_Filter.Add(new Dictionary<string, string>());
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Row", Convert.ToString(intCurrentJournalItem));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_ID", "Mon-" + Convert.ToString(intCurrentBypassPort));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Port_A_Name", Convert.ToString(strCurrentPortName));
                        list_CableJournal_Bypass_Filter[list_CableJournal_Bypass_Filter.Count - 1].Add("Device_A_Name", strCurrentDeviceHostname);


                        doubNextPortStartPointY -= 0.7;
                    };

                };

                doubBypassEndX = doubStartPointNextShapeX + 1;

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~







                //Console.WriteLine($"Total Bypass Hydra Count: {listBypassHydraConnectors.Count}");







                doubBypassEndX = doubStartPointNextShapeX + 1;



                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~







                int intJournalCurrentRow = 0;



















                int intLastUplinkPortOnBalancer = 0;
                int intUsedUplinkHydrasCounter = 0;

                int intLastDownlinkPortOnBalancer = 0;
                int intUsedDownlinkHydrasCounter = 0;


                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                Console.WriteLine($"Отрисовка байпасов завершена.");



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

                Visio.Shape[] arrShapesBalancerDevices = new Visio.Shape[10];
                Visio.Shape[] arrShapesFilterDevices = new Visio.Shape[100];

                doubStartPointNextShapeX += 100;


                intdoubTopLineY = doubUpperStartPoint + 2;



                doubNextPortStartPointX = doubStartPointNextShapeX;
                doubNextPortStartPointY = doubStartPointNextShapeY;

                int intGapBwBalancers = 10;
                int intFilterPointerCross;
                int intFilterPointerStraight = 0;
                int intHydraStartPoint;

                int intHydraPointerCross;
                int intHydraPointerStraight = 0;

                //Вычисление общего количеств портов 10G к балансировщикам. Удвоенное количество операторских 10G-линков. Возможно, округлить в большую сторону
                int intTotalBalancerUplink100Ports = intLinkCounter100 * 2;
                int intBalancerUplinkPortPerDevice = 0;
                int intTotalFiltersQuantity = Convert.ToInt32(strFilterNumberFromInput);
                int intTotalBalancersQuantity = 0;
                int intHydrasQuantityBwOnePair = 0;
                int intSubseqHydrasCounter;
                int intCrossPortsOnEachBalancer = 0;


                int intTotalFilterHydrasQuantity = intHydrasOnFilter * intTotalFiltersQuantity; //strFilterNumberFromInput;

                if (!boolNoBalancer)
                {
                    Console.WriteLine($"Отрисовка балансировщиков начата.");
                    intTotalBalancersQuantity = Convert.ToInt32(strBalancerNumberFromInput);
                    intCrossPortsOnEachBalancer = CalculateDevicesQuantity(intTotalBalancerUplink100Ports, intTotalBalancersQuantity);
                    if (intCrossPortsOnEachBalancer % 2 > 0) intCrossPortsOnEachBalancer++;
                    intBalancerUplinkPortPerDevice = intTotalFilterHydrasQuantity / intTotalBalancersQuantity;
                    if (intTotalBalancerUplink100Ports % intTotalBalancersQuantity > 0) Console.WriteLine("Нацело не делится");
                    intHydrasQuantityBwOnePair = CalculateDevicesQuantity(intTotalFilterHydrasQuantity, intTotalBalancersQuantity * intTotalFiltersQuantity);
                };

                Console.WriteLine($"~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                Console.WriteLine($"Строим крестовую схему.");
                Console.WriteLine($"~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                Console.WriteLine($"Балансировщики: {intTotalBalancersQuantity}, Net-порты на каждом балансировщике: {intCrossPortsOnEachBalancer}.");
                //Console.WriteLine($"Балансировщики: {intTotalBalancersQuantity}, Фильтры: {intTotalFiltersQuantity}.");
                //Console.WriteLine($"Всего {intTotalFilterHydrasQuantity} гидр на {intTotalFiltersQuantity} фильтрах {strFilterModel}.");
                //Console.WriteLine($"С каждого балансировщика по {intBalancerUplinkPortPerDevice} гидр на фильтры. По {intHydrasQuantityBwOnePair} гидр на фильтр.");
                Console.WriteLine($"~~~~~~~~~~~~~~~~~~~~~~~~~~~~");

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw ELB Device ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                //Console.WriteLine($"На Балансировщиках:");

                doubStartPointNextShapeX += 10;
                doubStartPointNextShapeY += 10;

                strCurrentDeviceHostname = "";
                //Для варианта с балансировщиком. Коммутация линков 10G гидрами.
                if (!boolNoBalancer)
                {
                    //Console.WriteLine($"Балансировщик !!!");
                    for (int intCurrentBalancerFrame = 1; intCurrentBalancerFrame <= intTotalBalancersQuantity; intCurrentBalancerFrame++)
                    {
                        arrShapesBalancerDevices[intCurrentBalancerFrame] = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 9);
                        strCurrentDeviceHostname = "ELB-0133 (" + intCurrentBalancerFrame + ")";
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

                        intStartBalancerPort = 17;

                        //Console.WriteLine($"Всего линков 100G: {intTotalBalancerUplink100Ports}");

                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw ELB to IS100 Uplink Ports (New)   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        if (intTotalBalancerUplink100Ports > 0)
                        {
                            for (int intCurrentBalancerUplinkPort = intStartBalancerPort; intCurrentBalancerUplinkPort - 16 <= intCrossPortsOnEachBalancer; intCurrentBalancerUplinkPort++)
                            {
                                //Console.WriteLine($"Net-порт балансировщика {intCurrentBalancerFrame}: {intCurrentBalancerUplinkPort}");
                                strCurrentPortName = "p" + intCurrentBalancerUplinkPort;
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort].Data3 = Convert.ToString(intCurrentBalancerFrame);
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort].Text = strCurrentPortName;
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";
                                if (intCurrentBalancerUplinkPort % 2 == 0) doubNextPortStartPointY -= 0.4;
                                else doubNextPortStartPointY -= 0.2;
                                if (arrShapesBypass100MonPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort] != null)
                                {
                                    //Console.WriteLine($"Балансировщик {intCurrentBalancerFrame}, Порт {intCurrentBalancerUplinkPort} на байпасе существует.");
                                    arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort].AutoConnect(arrShapesBypass100MonPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                    //КЖ


                                    // Console.WriteLine($"Балансировщик: {intCurrentBalancerFrame}, Порт: {intCurrentBalancerUplinkPort}");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intCurrentBalancerUplinkPort, 0].Add("Device_B_Name", strCurrentDeviceHostname);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intCurrentBalancerUplinkPort, 0].Add("Port_B_Name", strCurrentPortName);
                                    list_CableJournal_Bypass_Balancer.Add(new Dictionary<string, string>());

                                    foreach (string key in arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intCurrentBalancerUplinkPort, 0].Keys)
                                    {
                                        list_CableJournal_Bypass_Balancer[list_CableJournal_Bypass_Balancer.Count - 1].Add(key, arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intCurrentBalancerUplinkPort, 0][key]);
                                    };

                                };




                            };
                            //intStartBalancerPort = intTotalBalancerUplink100Ports + 17;
                            intStartBalancerPort = 17;

                        };

                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw ELB to 10G Uplink Ports (IS40 & IBS1UP)   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        //Console.WriteLine($"Балансер: {intCurrentBalancerFrame}, Начальный порт балансера: {intStartBalancerPort}, Конечный порт балансера: {intLastUplinkPortOnBalancer}");

                        if (intBypassIs40HydrasTotal > 0)
                        {
                            //Console.WriteLine($"Check Balancer.");
                            Console.WriteLine($"intBypassIs40HydrasTotal - intUsedUplinkHydrasCounter = {intBypassIs40HydrasTotal} - {intUsedUplinkHydrasCounter} = {intBypassIs40HydrasTotal - intUsedUplinkHydrasCounter}.");

                            if (intBypassIs40HydrasTotal - intUsedUplinkHydrasCounter > 16) intLastUplinkPortOnBalancer = 32;
                            else intLastUplinkPortOnBalancer = 16 + intBypassIs40HydrasTotal - intUsedUplinkHydrasCounter;

                            for (int intCurrentBalancerUplinkPort = intStartBalancerPort; intCurrentBalancerUplinkPort <= intLastUplinkPortOnBalancer; intCurrentBalancerUplinkPort++)
                            {
                                //intLastUplinkPortOnBalancer++;
                                intUsedUplinkHydrasCounter++;
                                strCurrentPortName = "p" + intCurrentBalancerUplinkPort;
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort].Data3 = Convert.ToString(intCurrentBalancerFrame);
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort].Text = strCurrentPortName;
                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";
                                // if (intCurrentBalancerUplinkPort % 2 == 0) doubNextPortStartPointY -= 0.4;
                                // else doubNextPortStartPointY -= 0.2;
                                doubNextPortStartPointY -= 0.4;

                                arrShapesBalancer100UplinkPorts[intCurrentBalancerFrame, intCurrentBalancerUplinkPort].AutoConnect(arrBypassIs40HydraConnectors[intCurrentBalancerFrame, intCurrentBalancerUplinkPort], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                //Console.WriteLine($"Пройдено 1.");
                                //КЖ
                                for (int inCurrentHydraEnd = 1; inCurrentHydraEnd <= 4; inCurrentHydraEnd++)
                                {
                                    //Console.WriteLine($"Балансировщик {intCurrentBalancerFrame}, Порт {intCurrentBalancerUplinkPort}, Конец Гидры: {inCurrentHydraEnd}.");
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intCurrentBalancerUplinkPort, inCurrentHydraEnd].Add("Device_B_Name", strCurrentDeviceHostname);
                                    arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intCurrentBalancerUplinkPort, inCurrentHydraEnd].Add("Port_B_Name", strCurrentPortName + "-" + inCurrentHydraEnd);

                                    list_CableJournal_Bypass_Balancer.Add(new Dictionary<string, string>());

                                    foreach (string key in arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intCurrentBalancerUplinkPort, inCurrentHydraEnd].Keys)
                                    {
                                        list_CableJournal_Bypass_Balancer[list_CableJournal_Bypass_Balancer.Count - 1].Add(key, arr_CableJournal_Bypass_Balancer[intCurrentBalancerFrame, intCurrentBalancerUplinkPort, inCurrentHydraEnd][key]);
                                    };
                                };

                                //Console.WriteLine($"Добавлено в список");
                            };
                        };


                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw Downlink Ports (для прямой и крестовой)   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        intFilterPointerCross = 0;
                        //intHydraStartPoint = 2 * (intCurrentBalancerFrame - 1);
                        //intHydraStartPoint = 2 * (intHydrasQuantityBwOnePair - 1);
                        intHydraStartPoint = intHydrasQuantityBwOnePair * (intCurrentBalancerFrame - 1);
                        intHydraPointerCross = intHydraStartPoint;

                        Console.WriteLine($"Между одним балансером и одним фильтром {intHydrasQuantityBwOnePair} гидр.");

                        //~~~~~~~~~~~~~~~~~~~~~~~~  Крестовая   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        if (boolCrossLayout)
                        {
                            intSubseqHydrasCounter = 0;
                            //Console.WriteLine($"Балансер {intCurrentBalancerFrame}, Макс.портов: {intTotalFilterHydrasQuantity / intTotalBalancersQuantity}.");
                            for (int intCurrentBalancerDownlinkPort = 1; intCurrentBalancerDownlinkPort <= intTotalFilterHydrasQuantity / intTotalBalancersQuantity; intCurrentBalancerDownlinkPort++)
                            {
                                //Логика сдвига указателя фильтров и гидр одинакова.
                                strCurrentPortName = "p" + intCurrentBalancerDownlinkPort;
                                if (intCurrentBalancerDownlinkPort > 16) strCurrentPortName = "p" + (49 - intCurrentBalancerDownlinkPort);
                                if (intHydraPointerCross == intHydraStartPoint) intFilterPointerCross++;
                                intHydraPointerCross++;
                                intSubseqHydrasCounter++;



                                Console.WriteLine($"Балансер {intCurrentBalancerFrame}, Порт {intCurrentBalancerDownlinkPort} ------ Фильтр {intFilterPointerCross}, Гидра {intHydraPointerCross}.");
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerCross, intHydraPointerCross] = page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 0.5, doubNextPortStartPointY - 0.3);
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerCross, intHydraPointerCross].Data3 = Convert.ToString(intCurrentBalancerFrame);
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerCross, intHydraPointerCross].Text = strCurrentPortName;
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerCross, intHydraPointerCross].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";
                                doubNextPortStartPointY -= 0.4;


                                for (int intCurrentPortInHydra = 1; intCurrentPortInHydra <= 4; intCurrentPortInHydra++)
                                {
                                    //КЖ "Балансировщики - Фильтры"
                                    intGlobalCableCounter++;
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra] = new Dictionary<string, string>();
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бф");
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Cable_Name", "ELB-0133 --- " + strFilterModel);
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Port_A_Name", strCurrentPortName + "-" + intCurrentPortInHydra);
                                    arrCableJournal_Balancer_Filter[intFilterPointerCross, intHydraPointerCross, intCurrentPortInHydra].Add("Device_A_Name", strCurrentDeviceHostname);
                                };

                                // Если текущий номер гидры совпал с количеством стыков балансера на один фильтр, сдвигаем указатель в начало .
                                if (intSubseqHydrasCounter == intHydrasQuantityBwOnePair)
                                {
                                    intHydraPointerCross = intHydraStartPoint;
                                    intSubseqHydrasCounter = 0;
                                };

                                //else if (intHydraPointerCross == intHydrasOnFilter) intHydraPointerCross = intHydraStartPoint;

                            };
                        }
                        //~~~~~~~~~~~~~~~~~~~~~~~~  Прямая   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        else
                        {
                            if (intTotalFilterHydrasQuantity - intUsedDownlinkHydrasCounter > intHydrasOnFilter) intLastDownlinkPortOnBalancer = 16;
                            else intLastDownlinkPortOnBalancer = intTotalFilterHydrasQuantity - intUsedDownlinkHydrasCounter;
                            //Console.WriteLine($"intTotalFilterHydrasQuantity - intUsedDownlinkHydrasCounter = {intTotalFilterHydrasQuantity} - {intUsedDownlinkHydrasCounter} = {intTotalFilterHydrasQuantity - intUsedDownlinkHydrasCounter}.");
                            Console.WriteLine($"Балансер {intCurrentBalancerFrame}, Последний порт: {intLastDownlinkPortOnBalancer}.");

                            for (int intCurrentBalancerDownlinkPort = 1; intCurrentBalancerDownlinkPort <= intLastDownlinkPortOnBalancer; intCurrentBalancerDownlinkPort++)
                            {
                                //Console.Write($"Фильтр {intFilterPointerStraight}, Гидра {intHydraPointerStraight} ---> ");
                                strCurrentPortName = "p" + intCurrentBalancerDownlinkPort;
                                //if (intHydraPointerStraight == 0) Console.WriteLine($"Гидра 0! Балансер {intCurrentBalancerFrame}, Порт {intLastDownlinkPortOnBalancer} --- Фильтр {intFilterPointerStraight}.");
                                if (intHydraPointerStraight == 0) intFilterPointerStraight++;
                                intHydraPointerStraight++;

                                //Console.WriteLine($"Фильтр {intFilterPointerStraight}, Гидра {intHydraPointerStraight} --->");

                                Console.WriteLine($"Балансер {intCurrentBalancerFrame}, Порт {intCurrentBalancerDownlinkPort} ------ Фильтр {intFilterPointerStraight}, Гидра {intHydraPointerStraight}.");
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerStraight, intHydraPointerStraight] = page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 0.5, doubNextPortStartPointY - 0.3);
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerStraight, intHydraPointerStraight].Data3 = Convert.ToString(intCurrentBalancerFrame);
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerStraight, intHydraPointerStraight].Text = strCurrentPortName;
                                arrShapesBalancer1040DownlinkPorts[intFilterPointerStraight, intHydraPointerStraight].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";
                                doubNextPortStartPointY -= 0.4;

                                for (int intCurrentPortInHydra = 1; intCurrentPortInHydra <= 4; intCurrentPortInHydra++)
                                {
                                    //КЖ "Балагсировщики - Фильтры"
                                    intGlobalCableCounter++;
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra] = new Dictionary<string, string>();
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Row", Convert.ToString(intCurrentJournalItem));
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "бф");
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Cable_Name", "ELB-0133 --- " + strFilterModel);
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Port_A_Name", strCurrentPortName + "-" + intCurrentPortInHydra);
                                    arrCableJournal_Balancer_Filter[intFilterPointerStraight, intHydraPointerStraight, intCurrentPortInHydra].Add("Device_A_Name", strCurrentDeviceHostname);
                                };

                                if (intHydraPointerStraight == intHydrasOnFilter) intHydraPointerStraight = 0;

                            };



                        };


                        doubStartPointNextShapeY -= intGapBwBalancers;

                        doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                        doubNextPortStartPointY = doubStartPointNextShapeY;
                        //doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                        doubNextPortStartPointY = doubStartPointNextShapeY + intGapBwBalancers;

                    };
                };















                //= 14*E2 + B2 - 2*C2*(D2-1)
                /*
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw 40G Ports ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                foreach (Visio.Shape objBypassMonPort in listShapesBypass40MonPorts)
                {
                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Drawing a New Balancer   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    //intCurrentBalancerPort++;
                    if (intCurrentBalancerPort == 32)
                    {
                        //Console.WriteLine($"Check");

                        intCurrentBalancerDevice++;
                        intCurrentBalancerPort = 16;
                        listShapesBalancerDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 9));
                        strCurrentDeviceHostname = "ELB-0133-" + intCurrentBalancerDevice + "40G";
                        listShapesBalancerDevices[listShapesBalancerDevices.Count - 1].Text = strCurrentDeviceHostname;
                        listShapesBalancerDevices[listShapesBalancerDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(143,188,143)";

                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.15, doubStartPointNextShapeX + 0.8, doubStartPointNextShapeY + 0.35));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MGMT";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MGMT";


                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw Balancer Downlink Ports ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                        doubNextPortStartPointY = doubStartPointNextShapeY + 7;

                        for (int intCurrentFloodedBalancerPort = 1; intCurrentFloodedBalancerPort <= intTotalFiltersFromBw * intHydrasOnFilter; intCurrentFloodedBalancerPort++)
                        {
                            doubPortStartX = doubNextPortStartPointX;
                            doubPortStartY = doubNextPortStartPointY - 7 - 0.5;
                            doubPortEndX = doubNextPortStartPointX + 0.5;
                            doubPortEndY = doubNextPortStartPointY - 7 - 0.3;

                            listShapesBalancerDownlinkPorts.Add(page1.DrawRectangle(doubPortStartX, doubPortStartY, doubPortEndX, doubPortEndY));
                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].Data3 = Convert.ToString(intCurrentBalancerDevice);

                            strCurrentPortName = "p" + intCurrentFloodedBalancerPort;
                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].Text = strCurrentPortName;
                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                            for (int intCurrentPortInHydra = 1; intCurrentPortInHydra <= 4; intCurrentPortInHydra++)
                            {
                                intGlobalCableCounter++;
                                //Console.WriteLine($"Check2. HydraCount: {intCurrentPortInHydra}.");
                                listCableJournal_3.Add(new Dictionary<string, string>());
                                //Console.WriteLine($"Balancer Port Index: {listCableJournal_3.Count - 1}, Device: {strCurrentDeviceHostname}, Port: {strCurrentPortName}");
                                listCableJournal_3[listCableJournal_3.Count - 1].Add("Row", Convert.ToString(intCurrentJournalItem));
                                listCableJournal_3[listCableJournal_3.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter));
                                listCableJournal_3[listCableJournal_3.Count - 1].Add("Cable_Name", "ELB-0133 --- " + strFilterModel);
                                listCableJournal_3[listCableJournal_3.Count - 1].Add("Port_A_Name", Convert.ToString(strCurrentPortName) + " (Гидра)");
                                listCableJournal_3[listCableJournal_3.Count - 1].Add("Device_A_Name", strCurrentDeviceHostname);

                            };

                            doubNextPortStartPointY -= 0.5;

                        };


                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;
                        doubStartPointNextShapeY -= 7;

                    };






                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Drawing 40G Ports on Balancer   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    intCurrentBalancerPort++;
                    //if (intCurrentBalancerPort == 33) continue;

                    //Console.Write($"Port Counter: {intCurrentBalancerPort}, Calculate: {intCurrentBalancerPort % 2}, ");
                    if (intCurrentBalancerPort % 2 == 0) intPortNumberAfterSwap = intCurrentBalancerPort - 1;
                    else intPortNumberAfterSwap = intCurrentBalancerPort + 1;
                    // Console.WriteLine($"Port No After Swap: {intPortNumberAfterSwap}");
                    //objBypassNetPort.Data2

                    if (intPortNumberAfterSwap % 2 == 0) doubShiftY = 0.2;
                    else doubShiftY = 0;


                    listShapesBalancerUplinkPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5 - doubShiftY, doubNextPortStartPointX, doubNextPortStartPointY - 0.3 - doubShiftY));
                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Data1 = Convert.ToString(intCurrentBalancerDevice);
                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Data2 = objBypassMonPort.Data2;

                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Text = "p" + intPortNumberAfterSwap;

                    objBypassMonPort.AutoConnect(listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1], Visio.VisAutoConnectDir.visAutoConnectDirNone);






                    if (intPortNumberAfterSwap % 2 > 0) doubNextPortStartPointY -= 0.5;
                    //else doubNextPortStartPointY += 0.8;

                };


                */

                //double doubHydraStartX;



                //Console.Write($"CurrentBalancerPorts: {intCurrentBalancerPort}, Calculate: {intCurrentBalancerPort % 2}, ");




                //if (intTotalIs40Bypasses > 0 | (intTotalIs10Bypasses > 0) && !boolEolBypass) Console.WriteLine($"IS40 Number: {intTotalIs40Bypasses} + {intTotalIs10Bypasses} = {intTotalIs40Bypasses + intTotalIs10Bypasses}");
                //if ((intTotalIs10Bypasses > 0) && boolEolBypass) Console.WriteLine($"IBS1U Number: {intTotalIs10Bypasses}");





                //Рисунок ODF 1


                Visio.Documents visioDocs = app.Documents;
                Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss", (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);



                //Visio.Shape shapeTest1 = page1.DrawRectangle(doubStartPointNextShapeX - 10, doubStartPointNextShapeY + 40, doubStartPointNextShapeX - 5, doubStartPointNextShapeY);

                //Visio.Document currentStencil = Application.

                //Visio.Shape shapeTest3 = page1.DrawRectangle(doubStartPointNextShapeX - 10, doubStartPointNextShapeY + 40, doubStartPointNextShapeX - 5, doubStartPointNextShapeY);
                //Visio.Shape shapeTest2 = page1.DrawRectangle(doubStartPointNextShapeX - 20, doubStartPointNextShapeY + 20, doubStartPointNextShapeX - 15, doubStartPointNextShapeY - 20);


                //Visio.Master currentStencil = doc.Masters.get_ItemU(@"Manager Belt");

                // Load the stencil we want
                //Visio.Master vsoTestStencil = doc.Masters.get_ItemU("Basic_U.vss");
                //Visio.Master vsoTestStencil = doc.Masters.get_ItemU(@"Manager Belt");

                //Visio.Master vsoTestStencil = doc.Masters.

                // show the stencil window
                //Visio.Window stencilWindow = page1.Document.OpenStencilWindow();

                // create a triangle shape from the stencil master collection
                //Visio.Shape shape1 = page1.Drop(vsoTestStencil, 1, 1);


                // this gives a count of all the stencils on the status bar
                //int countStencils = vsoTestStencil.;




                //Visio.Master[] masterShapes = new Visio.Master[2];

                //masterShapes[0] = doc.Masters.get_ItemU(@"Manager Belt");
                //masterShapes[1] = doc.Masters.get_ItemU(@"Vacancy Belt");





                // Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                //Console.WriteLine($"Балансировщики: {intTempCurrentBalancerDevice}");
                //Console.WriteLine($"Порты на фильтры: {listShapesBalancerDownlinkPorts.Count}");
                int intFiltersNumberFromPorts = listShapesBalancerDownlinkPorts.Count / 4;

                int intTotalFiltersFinal = Math.Max(intTotalFiltersFromBw, intFiltersNumberFromPorts);
                //if (intTotalFiltersFinal > intTempCurrentBalancerDevice * 4) intTotalFiltersFinal = intTempCurrentBalancerDevice * 4;

                intTotalFiltersFinal = intTotalFiltersFromBw;

                //Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                //Console.WriteLine($"Итоговое количество фильтров: {intTotalFiltersFinal}");

                int intMaximumHydras = intTotalFiltersFinal * 4;

                //Console.WriteLine($"Количество гидр 'балансировщики-фильтры': {intMaximumHydras}");
                //Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                /// Draw All Downlink Ports
                /// 

                //listShapesBalancerDownlinkPorts


                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////////////////////////////////////  Draw Filters ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                //intTotalFiltersFromBw


                List<Visio.Shape> listDeviceLogPorts = new List<Visio.Shape>();                            //Массив для LOG-портов фильтров


                //int intDownlinkPortOnBalancerToConnectFilterHydra;

                doubStartPointNextShapeX += 15;

                doubStartPointNextShapeY = doubUpperStartPoint;



                List<Visio.Shape> listShapesFilter4160Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesFilterPorts = new List<Visio.Shape>();
                List<Visio.Shape> listFilterHydraConnectors = new List<Visio.Shape>();


                List<Visio.Shape> listShapesSingleFilterPorts = new List<Visio.Shape>();


                // if (boolNoBalancer) Console.WriteLine("No Balancer");

                //int intBalancerJournalRow;

                int intCurrentHydraOnFilter;

                //int intHydraToConnect;


                //int intTotalFiltersQuantity = Convert.ToInt32(strFilterNumberFromInput);
                //int intTotalBalancersQuantity = Convert.ToInt32(strBalancerNumberFromInput);

                //Console.WriteLine($"Портов на балансере: {listShapesBalancerDownlinkPorts.Count}");

                // Для крестовой схемы 100G
                //if (boolCrossLayout)
                if (!boolNoBalancer)
                {

                    //for (int intCurrentFilterFrame = 1; intCurrentFilterFrame <= intTotalFiltersFinal; intCurrentFilterFrame++)
                    for (int intCurrentFilterFrame = 1; intCurrentFilterFrame <= intTotalFiltersQuantity; intCurrentFilterFrame++)
                    {


                        //--------------------------------------------------  Old Filter Chassis  ---------------------------------------------
                        /*
                        listShapesFilter4160Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 6));
                        strCurrentDeviceHostname = "Filter-" + strFilterModel + " (" + intCurrentFilterFrame + ")";
                        listShapesFilter4160Devices[listShapesFilter4160Devices.Count - 1].Text = strCurrentDeviceHostname;
                        listShapesFilter4160Devices[listShapesFilter4160Devices.Count - 1].Data3 = strCurrentDeviceHostname;
                        listShapesFilter4160Devices[listShapesFilter4160Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(175,238,238)";
                        */

                        //--------------------------------------------------  New Filter Chassis Start ---------------------------------------------

                        //doubStartPointNextShapeX += 10;
                        //doubStartPointNextShapeY += 10;

                        arrShapesFilterDevices[intCurrentFilterFrame] = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 6);
                        strCurrentDeviceHostname = "Filter-" + strFilterModel + " (" + intCurrentFilterFrame + ")";
                        arrShapesFilterDevices[intCurrentFilterFrame].Text = strCurrentDeviceHostname;
                        arrShapesFilterDevices[intCurrentFilterFrame].Data3 = strCurrentDeviceHostname;
                        arrShapesFilterDevices[intCurrentFilterFrame].get_Cells("FillForegnd").FormulaU = "=RGB(175,238,238)";

                        intCurrentHydraOnFilter = 0;


                        //---------------------  New Filter Ports Start ----------------------------------------

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        //if (boolCrossLayout)
                        // Логика определения портов для прмой и крестовой схем производится на стороне балансировщиков.
                        for (int intCurrentFilterPort = 1; intCurrentFilterPort <= 4 * intHydrasOnFilter; intCurrentFilterPort++)
                        {
                            strCurrentPortName = "p" + intCurrentFilterPort;
                            arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3);
                            arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort].Data3 = Convert.ToString(intCurrentFilterFrame);
                            arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort].Text = strCurrentPortName;
                            arrShapesFilterPorts[intCurrentFilterFrame, intCurrentFilterPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                            listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX - 0.7, doubNextPortStartPointY - 0.4, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4));
                            if (listHydraLines.Count == 1) intCurrentHydraOnFilter++;


                            //Console.WriteLine($"Фильтр {intCurrentFilterFrame}, Гидра {intCurrentHydraOnFilter}, Подпорт {listHydraLines.Count}.");

                            // 
                            arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, listHydraLines.Count].Add("Port_B_Name", strCurrentPortName + " (AOC c" + listHydraLines.Count + ")");
                            arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, listHydraLines.Count].Add("Device_B_Name", strCurrentDeviceHostname);
                            list_CableJournal_Balancer_Filter.Add(new Dictionary<string, string>());
                            foreach (string key in arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, listHydraLines.Count].Keys)
                            {
                                list_CableJournal_Balancer_Filter[list_CableJournal_Balancer_Filter.Count - 1].Add(key, arrCableJournal_Balancer_Filter[intCurrentFilterFrame, intCurrentHydraOnFilter, listHydraLines.Count][key]);
                            };

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
                                arrShapesBalancer1040DownlinkPorts[intCurrentFilterFrame, intCurrentHydraOnFilter].AutoConnect(arrFilterHydraConnectors[intCurrentFilterFrame, intCurrentHydraOnFilter], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            };



                            if (intCurrentFilterPort % 2 > 0) doubNextPortStartPointY -= 0.2;
                            else doubNextPortStartPointY -= 0.5;
                        }

                        //---------------------  New Filter Ports End ----------------------------------------

                        //doubStartPointNextShapeX -= 10;
                        //doubStartPointNextShapeY -= 10;

                        //--------------------------------------------------  New Filter Chassis End ---------------------------------------------

                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MNG";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MNG";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 0.3));
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "IPMI";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "IPMI";
                        listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = strCurrentDeviceHostname;
                        listDeviceLogPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 1.1, doubStartPointNextShapeY + 0.3));
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Text = "SP1";
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Rotate90();
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data2 = "SP1";
                        listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data3 = strCurrentDeviceHostname;

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        //int intRecalculatedFilterPort;

                        /*
                        for (int intCurrentFilterPort = 1; intCurrentFilterPort <= 4 * intHydrasOnFilter; intCurrentFilterPort++)
                        {
                            listShapesFilterPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX, doubNextPortStartPointY - 0.2));
                            strCurrentPortName = "Te " + intCurrentFilterPort;
                            listShapesFilterPorts[listShapesFilterPorts.Count - 1].Text = strCurrentPortName;
                            listShapesFilterPorts[listShapesFilterPorts.Count - 1].Data1 = Convert.ToString(intCurrentFilterFrame);
                            listShapesFilterPorts[listShapesFilterPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                            ///////////////     Draw Hydra Connector    //////////////////////////

                            listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX - 0.7, doubNextPortStartPointY - 0.3, doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3));

                            if (listHydraLines.Count == 4)
                            {
                                listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX - 0.7, doubNextPortStartPointY - 0.3, doubNextPortStartPointX - 0.7, doubNextPortStartPointY + 0.6));
                                vsoWindow.DeselectAll();
                                foreach (Visio.Shape objHydraSingleLine in listHydraLines)
                                {
                                    vsoWindow.Select(objHydraSingleLine, 2);
                                };

                                vsoSelection = vsoWindow.Selection;


                                listFilterHydraConnectors.Add(vsoSelection.Group());
                                listFilterHydraConnectors[listFilterHydraConnectors.Count - 1].Data1 = Convert.ToString(intCurrentFilterFrame);
                                listHydraLines.Clear();

                                // Formula //

                                //if (boolCrossLayout)
                                if (boolCrossLayout & (intBalancersFinalQuantity > 1)) intDownlinkPortOnBalancerToConnectFilterHydra = intTotalFiltersFinal * (listFilterHydraConnectors.Count - (intCurrentFilterFrame - 1) * intHydrasOnFilter - 1) + intCurrentFilterFrame - 1;
                                else intDownlinkPortOnBalancerToConnectFilterHydra = listFilterHydraConnectors.Count - 1;


                                //Console.WriteLine($"Hydra: {listFilterHydraConnectors.Count}, Balancer Port: {intDownlinkPortOnBalancerToConnectFilterHydra}, intCurrentFilterFrame: {intCurrentFilterFrame}");


                                //listFilterHydraConnectors[listFilterHydraConnectors.Count - 1].AutoConnect(listShapesBalancerDownlinkPorts[intDownlinkPortOnBalancerToConnectFilterHydra], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                                //Console.WriteLine("Check");

                                //if (intCurrentSubslotCounterInBypass == 1) strCableInHydra = " (AOC c2)";
                                //else strCableInHydra = " (AOC c4)";
                                //if (boolNoBalancer) strCableInHydra = "";

                                for (int intCurrentPortInHydra = 1; intCurrentPortInHydra <= 4; intCurrentPortInHydra++)
                                {
                                    intRecalculatedFilterPort = intCurrentFilterPort - 4 + intCurrentPortInHydra;
                                    strCurrentPortName = "Te" + intRecalculatedFilterPort;
                                    intBalancerJournalRow = intDownlinkPortOnBalancerToConnectFilterHydra * 4 + intCurrentPortInHydra - 1;
                                    //Console.WriteLine($"Balancer Port: {intDownlinkPortOnBalancerToConnectFilterHydra}, Port in Hydra: {intCurrentPortInHydra}, Row: {intBalancerJournalRow}");
                                    //Console.WriteLine($"Filter Port Index: {intBalancerJournalRow}, Device: {strCurrentDeviceHostname}, Port: {strCurrentPortName}");
                                    listCableJournal_3[intBalancerJournalRow].Add("Device_B_Name", strCurrentDeviceHostname);
                                    listCableJournal_3[intBalancerJournalRow].Add("Port_B_Name", strCurrentPortName + " (AOC c" + intCurrentPortInHydra + ")");
                                };
                            };



                            if (intCurrentFilterPort % 2 > 0) doubNextPortStartPointY -= 0.2;
                            else doubNextPortStartPointY -= 0.5;

                        };

                        */


                        doubStartPointNextShapeY -= 7;


                    };
                }
                ////////////////////////////////    Working with Single-Filters without Balancers   //////////////////////////
                else
                {
                    /////////////////////////////////////   Define Single Filter Type   //////////////////////////////////////

                    string strSingleFilterType = "";
                    int intPortsNumberOnSingleFilter = 0;

                    switch (listShapesLanPorts.Count)
                    {
                        case int n when (n > 6 && n <= 8):
                            strSingleFilterType = "Ecofilter 4160";
                            intPortsNumberOnSingleFilter = 4;
                            //Console.WriteLine($"I am 4160. {listShapesLanPorts.Count} links.");
                            break;

                        case int n when (n <= 6 && n >= 5):
                            strSingleFilterType = "Ecofilter 4120";
                            intPortsNumberOnSingleFilter = 6;
                            // Console.WriteLine($"I am 4120. {listShapesLanPorts.Count} links.");
                            break;

                        case int n when (n <= 4):
                            strSingleFilterType = "Ecofilter 4080";
                            intPortsNumberOnSingleFilter = 8;
                            //Console.WriteLine($"I am 4080. {listShapesLanPorts.Count} links.");
                            break;
                    }

                    //Drawing Single-Balancer Filter

                    doubStartPointNextShapeX = doubBypassEndX + 4;
                    doubStartPointNextShapeY -= 1;

                    Visio.Shape objStandAloneFilter = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 0.4 * intPortsNumberOnSingleFilter);
                    objStandAloneFilter.Text = strSingleFilterType;
                    strCurrentDeviceHostname = strSingleFilterType;
                    objStandAloneFilter.get_Cells("FillForegnd").FormulaU = "=RGB(175,238,238)";

                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.3, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "MNG";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "MNG";
                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.5, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 0.9, doubStartPointNextShapeY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "IPMI";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Rotate90();
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "IPMI";
                    listDeviceLogPorts.Add(page1.DrawRectangle(doubStartPointNextShapeX + 0.7, doubStartPointNextShapeY + 0.1, doubStartPointNextShapeX + 1.1, doubStartPointNextShapeY + 0.3));
                    listDeviceLogPorts[listDeviceLogPorts.Count - 1].Text = "SP1";
                    listDeviceLogPorts[listDeviceLogPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceLogPorts[listDeviceLogPorts.Count - 1].Rotate90();
                    listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceLogPorts[listDeviceLogPorts.Count - 1].Data2 = "SP1";


                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    for (int intCurrentFilterPort = 1; intCurrentFilterPort <= intPortsNumberOnSingleFilter; intCurrentFilterPort++)
                    {
                        strCurrentPortName = "Te " + intCurrentFilterPort;

                        //listShapesFilterPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1));
                        //listShapesFilterPorts[listShapesFilterPorts.Count - 1].Text = strCurrentPortName;
                        //listShapesFilterPorts[listShapesFilterPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                        arrShapesFilterPorts[0, intCurrentFilterPort] = page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1);
                        arrShapesFilterPorts[0, intCurrentFilterPort].Text = strCurrentPortName;
                        arrShapesFilterPorts[0, intCurrentFilterPort].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                        arrShapesFilterPorts[0, intCurrentFilterPort].AutoConnect(arrShapesBypass10_MonPorts[0, intCurrentFilterPort, 0], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                        if (intCurrentFilterPort % 2 > 0) doubNextPortStartPointY -= 0.2;
                        else doubNextPortStartPointY -= 0.5;



                        //  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        //КЖ к байпасам
                        Console.WriteLine($"Фильтр Единственный, Порт {intCurrentFilterPort}");
                        arr_CableJournal_Bypass_Balancer[0, intCurrentFilterPort, 0].Add("Device_B_Name", strCurrentDeviceHostname);
                        arr_CableJournal_Bypass_Balancer[0, intCurrentFilterPort, 0].Add("Port_B_Name", strCurrentPortName);
                        list_CableJournal_Bypass_Balancer.Add(new Dictionary<string, string>());
                        foreach (string key in arr_CableJournal_Bypass_Balancer[0, intCurrentFilterPort, 0].Keys)
                        {
                            list_CableJournal_Bypass_Balancer[list_CableJournal_Bypass_Balancer.Count - 1].Add(key, arr_CableJournal_Bypass_Balancer[0, intCurrentFilterPort, 0][key]);
                        };

                    };

                    /*
                    //////  Group Standalone Filter & Ports ////////

                    vsoWindow.DeselectAll();
                    vsoWindow.Select(objStandAloneFilter, 2);
                    foreach (Visio.Shape objFilterPort in listShapesFilterPorts)
                    {
                        vsoWindow.Select(objFilterPort, 2);
                    };
                    vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1], 2);
                    vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 2], 2);
                    vsoWindow.Select(listDeviceLogPorts[listDeviceLogPorts.Count - 1], 2);
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                    */
                };



                intGlobalCableCounter = 0;

                //////////  Draw Continent Router   //////////


                doubStartPointNextShapeX = doubNextPortStartPointX;
                doubStartPointNextShapeY = intdoubTopLineY + 1;
                double doubContStartX = doubStartPointNextShapeX;
                double doubContStartY = doubStartPointNextShapeY + 0.8;

                string strContinentUplinkPort;
                string strContinentToMes3348Port;
                string strContinentToMes5332aPort;
                string strContinentHostname;

                if (boolContinentIpcR300)
                {
                    strContinentHostname = "IPC-R300";
                    strContinentUplinkPort = "ix2";             //0
                    strContinentToMes5332aPort = "ix3";         //1
                    strContinentToMes3348Port = "igb0";         //2
                }
                else
                {
                    strContinentHostname = "IPC-100";
                    strContinentUplinkPort = "0 или 2";         //0
                    strContinentToMes5332aPort = "1";           //1
                    strContinentToMes3348Port = "3";            //2
                };

                //Рисуем шасси континента
                Visio.Shape objContinentSwitch = page1.DrawRectangle(doubStartPointNextShapeX - 1, doubStartPointNextShapeY + 0.5, doubStartPointNextShapeX, doubStartPointNextShapeY + 2.2);
                objContinentSwitch.Text = strContinentHostname;
                objContinentSwitch.get_Cells("FillForegnd").FormulaU = "=RGB(176,196,222)";

                List<Visio.Shape> listShapesContinentPorts = new List<Visio.Shape>();
                doubNextPortStartPointX = doubStartPointNextShapeX - 3;
                doubNextPortStartPointY = doubStartPointNextShapeY + 0.2;

                //Рисуем аплинк-порт континента
                Visio.Shape visContinentUplinkPort = page1.DrawRectangle(doubStartPointNextShapeX - 1.4, doubStartPointNextShapeY + 0.7, doubStartPointNextShapeX - 1, doubStartPointNextShapeY + 0.9);
                visContinentUplinkPort.Text = strContinentUplinkPort;
                visContinentUplinkPort.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                //Рисуем порт континента к MES5332A
                Visio.Shape visContinentToMes5332aPort = page1.DrawRectangle(doubStartPointNextShapeX - 0.4, doubStartPointNextShapeY + 2.3, doubStartPointNextShapeX, doubStartPointNextShapeY + 2.5);
                visContinentToMes5332aPort.Text = strContinentToMes5332aPort;
                visContinentToMes5332aPort.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                visContinentToMes5332aPort.Rotate90();


                //Рисуем порт континента к MES3348
                Visio.Shape visContinentToMes3348Port = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY + 0.7, doubStartPointNextShapeX + 0.4, doubStartPointNextShapeY + 0.9);
                visContinentToMes3348Port.Text = strContinentToMes3348Port;
                visContinentToMes3348Port.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                doubNextPortStartPointY = doubStartPointNextShapeY + 3.3;

                //////////  Draw MES Log Switch   //////////

                Visio.Shape shapeMesLogDevice = page1.DrawRectangle(doubStartPointNextShapeX + 2, doubStartPointNextShapeY + 3.6, doubStartPointNextShapeX + 2 + listDeviceLogPorts.Count * 0.3, doubStartPointNextShapeY + 4.6 + intTotalLogServers * 0.1);
                shapeMesLogDevice.Text = "MES5332A (log)";
                shapeMesLogDevice.get_Cells("FillForegnd").FormulaU = "=RGB(176,196,222)";


                List<Visio.Shape> listShapesMesLogDownLinkPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesLogUpLinkPorts = new List<Visio.Shape>();

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
                    listShapesMesLogDownLinkPorts[listShapesMesLogDownLinkPorts.Count - 1].Data1 = "MES5332A";
                    listShapesMesLogDownLinkPorts[listShapesMesLogDownLinkPorts.Count - 1].Data2 = "XGE-" + intMes32Port;
                };


                //Рисуем на MES5332A порт GE-32 - на стык с континентом
                Visio.Shape visMes5332aToContinentPort = page1.DrawRectangle(doubStartPointNextShapeX + 1.6, doubStartPointNextShapeY + 4.3, doubStartPointNextShapeX + 2, doubStartPointNextShapeY + 4.5);
                visMes5332aToContinentPort.Text = "GE-32";
                visMes5332aToContinentPort.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                visMes5332aToContinentPort.AutoConnect(visContinentToMes5332aPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);

                //Добавляем в КЖ запись "5332А - Континент" - сразу обе стороны.
                intGlobalCableCounter++;
                listCableJournal_Management.Add(new Dictionary<string, string>());
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "c");
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", "MES5332A --- " + strContinentHostname);
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", "MES5332A");
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", "XGE-" + intMes32Port);
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", strContinentHostname);
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", strContinentToMes5332aPort);
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "LC-LC/UPC MM (50/125мкм) OM4 (3.0мм) 2м");

                doubNextPortStartPointX = doubStartPointNextShapeX + 2 + listDeviceLogPorts.Count * 0.3;
                doubNextPortStartPointY = doubStartPointNextShapeY + 3.9;


                //Рисуем на MES5332A порт OOB - на стык с последним в цепи MES3348
                Visio.Shape visMesLogPortToMgmt = page1.DrawRectangle(doubStartPointNextShapeX + 1.6, doubStartPointNextShapeY + 3.9, doubStartPointNextShapeX + 2, doubStartPointNextShapeY + 4.1);
                visMesLogPortToMgmt.Text = "OOB";
                visMesLogPortToMgmt.get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";



                //Рисуем на 5332А порты для стыков с серверами
                for (int intCurrentServer = 0; intCurrentServer <= intTotalLogServers; intCurrentServer++)
                {
                    listShapesMesLogUpLinkPorts.Add(page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.4, doubNextPortStartPointY + 0.2));
                    listShapesMesLogUpLinkPorts[listShapesMesLogUpLinkPorts.Count - 1].Text = "XGE-" + (30 - intCurrentServer);
                    listShapesMesLogUpLinkPorts[listShapesMesLogUpLinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listShapesMesLogUpLinkPorts[listShapesMesLogUpLinkPorts.Count - 1].Data1 = "MES5332A";
                    listShapesMesLogUpLinkPorts[listShapesMesLogUpLinkPorts.Count - 1].Data2 = "XGE-" + (30 - intCurrentServer);
                    doubNextPortStartPointY += 0.2;
                };








                // Сервер СПХД

                //double doubServersStartX = doubStartPointNextShapeX + 4 + listDeviceLogPorts.Count * 0.3;
                double doubServersStartX = doubStartPointNextShapeX + 8;
                double doubServersStartY = doubStartPointNextShapeY + 3;

                Visio.Shape shapeServerShdDevice = page1.DrawRectangle(doubServersStartX, doubServersStartY, doubServersStartX + 2, doubServersStartY + 0.6);
                //Visio.Shape shapeServerShdDevice = page1.DrawRectangle(doubStartPointNextShapeX + 8, doubStartPointNextShapeY + 3, doubStartPointNextShapeX + 10, doubStartPointNextShapeY + 3.6);

                shapeServerShdDevice.Text = "Сервер СПХД";
                shapeServerShdDevice.get_Cells("FillForegnd").FormulaU = "=RGB(210,180,140)";



                List<Visio.Shape> listShapesSrvPortsToLog = new List<Visio.Shape>();

                listShapesSrvPortsToLog.Add(page1.DrawRectangle(doubServersStartX - 0.5, doubStartPointNextShapeY + 3.1, doubServersStartX, doubStartPointNextShapeY + 3.3));
                listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].Text = "XGE-1 (верхний)";
                listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].AutoConnect(listShapesMesLogUpLinkPorts[0], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                intGlobalCableCounter++;
                listCableJournal_Management.Add(new Dictionary<string, string>());
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "c");
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", "СПХД --- " + listShapesMesLogUpLinkPorts[0].Data1);
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", "Сервер СПХД");
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", "XGE-1 (верхний)");
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", listShapesMesLogUpLinkPorts[0].Data1);
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", listShapesMesLogUpLinkPorts[0].Data2);
                listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "FT-SFP+CabA-2 (10G, SFP+, AOC, 2м)");

                listDeviceMgmtPorts.Add(page1.DrawRectangle(doubServersStartX + 2, doubStartPointNextShapeY + 3.1, doubServersStartX + 2.4, doubStartPointNextShapeY + 3.3));
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "GE-4";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = "Сервер СПХД";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "GE-4";

                listDeviceMgmtPorts.Add(page1.DrawRectangle(doubServersStartX + 2, doubStartPointNextShapeY + 3.3, doubServersStartX + 2.4, doubStartPointNextShapeY + 3.5));
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "Mgmt";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = "Сервер СПХД";
                listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "Mgmt";

                vsoWindow.DeselectAll();
                vsoWindow.Select(shapeServerShdDevice, 2);
                vsoWindow.Select(listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1], 2);
                vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1], 2);
                vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 2], 2);
                vsoSelection = vsoWindow.Selection;
                vsoSelection.Group();


                //Серверы СПФС

                List<Visio.Shape> listShapesSpfsDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesSpfsPorts = new List<Visio.Shape>();

                double doubSfsdDeviceX = doubStartPointNextShapeX + 8;
                double doubSfsdDeviceY = doubStartPointNextShapeY + 4;

                for (int intCurrentSpfsDevice = 1; intCurrentSpfsDevice <= intTotalLogServers; intCurrentSpfsDevice++)
                {
                    listShapesSpfsDevices.Add(page1.DrawRectangle(doubSfsdDeviceX, doubSfsdDeviceY, doubSfsdDeviceX + 2, doubSfsdDeviceY + 0.6));
                    strCurrentDeviceHostname = "Сервер СПФС (" + intCurrentSpfsDevice + ")";
                    listShapesSpfsDevices[listShapesSpfsDevices.Count - 1].Text = strCurrentDeviceHostname;
                    listShapesSpfsDevices[listShapesSpfsDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(210,180,140)";

                    listShapesSrvPortsToLog.Add(page1.DrawRectangle(doubSfsdDeviceX - 0.5, doubSfsdDeviceY + 0.1, doubSfsdDeviceX, doubSfsdDeviceY + 0.3));
                    listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].Text = "XGE-1 (верхний)";
                    listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1].AutoConnect(listShapesMesLogUpLinkPorts[listShapesSrvPortsToLog.Count - 1], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                    //listShapesMesLogUpLinkPorts
                    intGlobalCableCounter++;
                    listCableJournal_Management.Add(new Dictionary<string, string>());
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "c");
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", strCurrentDeviceHostname + " --- " + listShapesMesLogUpLinkPorts[listShapesSrvPortsToLog.Count - 1].Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", strCurrentDeviceHostname);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", "XGE-1 (верхний)");
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", listShapesMesLogUpLinkPorts[listShapesSrvPortsToLog.Count - 1].Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", listShapesMesLogUpLinkPorts[listShapesSrvPortsToLog.Count - 1].Data2);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "FT-SFP+CabA-2 (10G, SFP+, AOC, 2м)");

                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubSfsdDeviceX + 2, doubSfsdDeviceY + 0.1, doubSfsdDeviceX + 2.4, doubSfsdDeviceY + 0.3));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "GE-4";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "GE-4";



                    listDeviceMgmtPorts.Add(page1.DrawRectangle(doubSfsdDeviceX + 2, doubSfsdDeviceY + 0.3, doubSfsdDeviceX + 2.4, doubSfsdDeviceY + 0.5));
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Text = "Mgmt";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data1 = strCurrentDeviceHostname;
                    listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data2 = "Mgmt";

                    vsoWindow.DeselectAll();
                    vsoWindow.Select(listShapesSpfsDevices[listShapesSpfsDevices.Count - 1], 2);
                    vsoWindow.Select(listShapesSrvPortsToLog[listShapesSrvPortsToLog.Count - 1], 2);
                    vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1], 2);
                    vsoWindow.Select(listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 2], 2);
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();

                    doubSfsdDeviceY += 0.8;
                }

                //3348
                //////////  Draw MES Management Switches   //////////

                int intMes3348ChassisNumber = CalculateDevicesQuantity(listDeviceMgmtPorts.Count, 48);
                doubStartPointNextShapeX = doubNextPortStartPointX + 2.5;



                string strPrevious3348Hostname = "";
                string strPrevious3348Portname = "";
                strCurrentDeviceHostname = "";

                List<Visio.Shape> listShapesMesMgmtDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesMgmtPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesMesUplinkPorts = new List<Visio.Shape>();
                Visio.Shape[,] arrMes3348UplinkPorts = new Visio.Shape[10, 2];

                int intMes48Port = 46;
                int intMess3348CurrentChassis = 0;
                double doubSwitchWidth;
                if (listDeviceMgmtPorts.Count <= 46) doubSwitchWidth = listDeviceMgmtPorts.Count * 0.22;
                else doubSwitchWidth = 46 * 0.22;

                for (int intCurrentMesPort = 1; intCurrentMesPort <= listDeviceMgmtPorts.Count; intCurrentMesPort++)    // Не считать порты 47 и 48 в общем списке.
                {

                    // Рисуем новое шасси, если текущий порт = 1, 47 и т.д.
                    if (intCurrentMesPort % 46 == 1)
                    //////////////////////  Add New MES Switch  //////////////////////////
                    {
                        doubNextPortStartPointX += 1;
                        listShapesMesMgmtDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + doubSwitchWidth, doubStartPointNextShapeY + 1));
                        strCurrentDeviceHostname = "MES3348 (" + listShapesMesMgmtDevices.Count + ")";
                        listShapesMesMgmtDevices[listShapesMesMgmtDevices.Count - 1].Text = strCurrentDeviceHostname;
                        listShapesMesMgmtDevices[listShapesMesMgmtDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(176,196,222)";
                        listShapesMesMgmtDevices[listShapesMesMgmtDevices.Count - 1].Data3 = strCurrentDeviceHostname;
                        intMes48Port = 0;
                        doubNextPortStartPointX = doubStartPointNextShapeX + 0.1;
                        doubNextPortStartPointY = doubStartPointNextShapeY - 0.3;

                        //На новом свитче рисуем три служебных порта

                        //Порт GE-48 вставляется слева на любом свитче.
                        intMess3348CurrentChassis++;
                        arrMes3348UplinkPorts[intMess3348CurrentChassis, 0] = page1.DrawRectangle(doubStartPointNextShapeX - 0.4, doubStartPointNextShapeY + 0.2, doubStartPointNextShapeX, doubStartPointNextShapeY + 0.4);
                        arrMes3348UplinkPorts[intMess3348CurrentChassis, 0].Text = "GE-48";
                        arrMes3348UplinkPorts[intMess3348CurrentChassis, 0].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";


                        //Если свитч первый, его порт GE-48 стыкуется с континентом.
                        if (intMess3348CurrentChassis == 1)
                        {
                            strPrevious3348Hostname = strContinentHostname;
                            strPrevious3348Portname = strContinentToMes3348Port;
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 0].AutoConnect(visContinentToMes3348Port, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            intGlobalCableCounter++;
                            listCableJournal_Management.Add(new Dictionary<string, string>());
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "c");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", strContinentHostname + " - " + strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", strContinentHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", strContinentToMes3348Port);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", "GE-48");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "UTP cat. 5e (RJ45-RJ45)");
                            doubContStartY += 0.2;
                        }

                        //Если свитч не первый, его порт GE-48 стыкуется с GE-47 предыдущего свитча.
                        else
                        {
                            //Console.WriteLine($"Начали соединять порт GE-47 MES3348-{intMess3348CurrentChassis}.");
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 0].AutoConnect(arrMes3348UplinkPorts[intMess3348CurrentChassis - 1, 1], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            //Console.WriteLine($"Провели линию.");
                            //удалить после свопа
                            strPrevious3348Portname = "GE-47";
                            intGlobalCableCounter++;
                            listCableJournal_Management.Add(new Dictionary<string, string>());
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "c");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", strPrevious3348Hostname + " --- " + strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", strPrevious3348Hostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", "GE-47");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", "GE-48");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "UTP cat. 5e (RJ45-RJ45)");

                            strPrevious3348Hostname = strContinentHostname;
                            strPrevious3348Portname = strContinentToMes3348Port;

                            //Console.WriteLine($"Соединили порт GE-47 MES3348-{intMess3348CurrentChassis}.");
                        };

                        //Если свитч последний, его порт GE-47 рисуется слева, стыковка с ООВ MES5332A.
                        if (intMess3348CurrentChassis == intMes3348ChassisNumber)
                        {
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1] = page1.DrawRectangle(doubStartPointNextShapeX - 0.4, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX, doubStartPointNextShapeY + 0.6);
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1].Text = "GE-47";
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1].AutoConnect(visMesLogPortToMgmt, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                            intGlobalCableCounter++;
                            listCableJournal_Management.Add(new Dictionary<string, string>());
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "c");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", "MES5332A --- " + strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", "MES5332A");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", "OOB");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", strCurrentDeviceHostname);
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", "GE-47");
                            listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "UTP cat. 5e (RJ45-RJ45)");
                        }
                        //Если свитч не последний, его порт GE-47 рисуется справа. Стык - при настройке следующего свитча.
                        else
                        {
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1] = page1.DrawRectangle(doubStartPointNextShapeX + doubSwitchWidth, doubStartPointNextShapeY + 0.4, doubStartPointNextShapeX + doubSwitchWidth + 0.4, doubStartPointNextShapeY + 0.6);
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1].Text = "GE-47";
                            arrMes3348UplinkPorts[intMess3348CurrentChassis, 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08"; ;
                        };



                        doubStartPointNextShapeX += 20;


                    };

                    //Console.WriteLine($"Соединяем порты MES3348-{intMess3348CurrentChassis}.");

                    //////////////////////  Add MES Ports  //////////////////////////

                    intMes48Port++;
                    listShapesMesMgmtPorts.Add(page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.4, doubNextPortStartPointY + 0.2));
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Text = "GE-" + intMes48Port;
                    doubNextPortStartPointX += 0.2;
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Rotate90();
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Data1 = "MES3348-" + listShapesMesMgmtDevices.Count;
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Data2 = "GE-" + intMes48Port;
                    listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Data3 = strCurrentDeviceHostname;
                    vsoWindow.Select(listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1], 2);

                };



                doubStartPointNextShapeX += listDeviceMgmtPorts.Count * 0.25 + 1;







                //vsoWindow.DeselectAll();

                // Не работает группирование на MES3348!!! Разобраться!!!

                //vsoWindow.DeselectAll();
                //vsoWindow.Select(listShapesMesMgmtDevices[listShapesMesMgmtDevices.Count - 1], 2);
                //vsoWindow.Select(listShapesMesUplinkPorts[listShapesMesUplinkPorts.Count - 1], 2);
                // foreach (Visio.Shape objMesMgmtPort in listShapesMesMgmtPorts)
                // {
                //     vsoWindow.Select(objMesMgmtPort, 2);
                // };
                //vsoSelection = vsoWindow.Selection;
                //vsoSelection.Group();


                int intMesPortCurrentPort = 0;





                foreach (Visio.Shape objDeviceMgmtPort in listDeviceMgmtPorts)
                {
                    objDeviceMgmtPort.AutoConnect(listShapesMesMgmtPorts[intMesPortCurrentPort], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                    intGlobalCableCounter++;
                    listCableJournal_Management.Add(new Dictionary<string, string>());
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "c");
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", objDeviceMgmtPort.Data1 + " --- " + listShapesMesMgmtPorts[intMesPortCurrentPort].Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", objDeviceMgmtPort.Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", objDeviceMgmtPort.Data2);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", listShapesMesMgmtPorts[intMesPortCurrentPort].Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", listShapesMesMgmtPorts[intMesPortCurrentPort].Data2);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "UTP cat. 5e (RJ45-RJ45)");

                    intMesPortCurrentPort++;

                };


                intMesPortCurrentPort = 0;


                //Console.WriteLine($"Логпортов: {listDeviceLogPorts.Count}, Кабелей до логпортов: {intGlobalCableCounter}");

                foreach (Visio.Shape objDeviceLogPort in listDeviceLogPorts)
                {
                    intGlobalCableCounter++;
                    objDeviceLogPort.AutoConnect(listShapesMesLogDownLinkPorts[intMesPortCurrentPort], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                    // Console.WriteLine($"Кабель: {intGlobalCableCounter}, Имя: {objDeviceLogPort.Data1 + " --- " + listShapesMesLogDownLinkPorts[intMesPortCurrentPort].Data1}");
                    listCableJournal_Management.Add(new Dictionary<string, string>());
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Number", Convert.ToString(intGlobalCableCounter) + "c");
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Name", objDeviceLogPort.Data1 + " --- " + listShapesMesLogDownLinkPorts[intMesPortCurrentPort].Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_A", objDeviceLogPort.Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_A", objDeviceLogPort.Data2);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Device_B", listShapesMesLogDownLinkPorts[intMesPortCurrentPort].Data1);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Port_B", listShapesMesLogDownLinkPorts[intMesPortCurrentPort].Data2);
                    listCableJournal_Management[listCableJournal_Management.Count - 1].Add("Cable_Type", "FT-SFP+CabA-2 (10G, SFP+, AOC, 2м)");

                    intMesPortCurrentPort++;

                };



                //};




                //long[] In_connectedShapeIds = Visio.Shape.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                //long[] Out_connectedShapeIds = Visio.Shape.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesOutgoingNodes, null);


                //////  Group Continent & Ports ////////
                ///
                /*
                vsoWindow.DeselectAll();
                vsoWindow.Select(objContinentSwitch, 2);
                foreach (Visio.Shape objContinentPort in listShapesContinentPorts)
                {
                    vsoWindow.Select(objContinentPort, 2);
                };
                vsoSelection = vsoWindow.Selection;
                vsoSelection.Group();
                */

                //Visio.Shape Shape1 = page1.DrawRectangle(30, 30, 40, 40);
                //Visio.Shape Shape2 = page1.DrawRectangle(50, 50, 60, 60);

                //Visio.Shape Connect1 = page1.DropConnected(Shape1, Shape2, Visio.VisAutoConnectDir.visAutoConnectDirNone);

                //Visio.Shape Connect2 = page1.Drop(Shape1,70,70);

                //Connect1.Text = "Zhopa";

                //////  Group 5332A Ports ////////
                ///

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



                /*
                //////  Group 3348 Ports ////////
                ///

                foreach (Visio.Shape objSingle3348Device in listShapesMesMgmtDevices)
                {
                    vsoWindow.DeselectAll();
                    vsoWindow.Select(objSingle3348Device, 2);
    
                    foreach (Visio.Shape objSinglePort in listShapesMesUplinkPorts)
                    {
                        if (objSinglePort.Data3 == objSingle3348Device.Data3)
                        {
                            vsoWindow.Select(objSinglePort, 2);
                        };
                    };

                    foreach (Visio.Shape objSinglePort in listShapesMesMgmtPorts)
                    {
                        if (objSinglePort.Data3 == objSingle3348Device.Data3)
                        {
                            vsoWindow.Select(objSinglePort, 2);
                        };
                    };

                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };



            */


                //vsoWindow.DeselectAll();
                //listShapesMesMgmtPorts[listShapesMesMgmtPorts.Count - 1].Data3 = Convert.ToString(listShapesMesMgmtDevices.Count);
                //listShapesMesMgmtDevices[listShapesMesMgmtDevices.Count - 1].Data3 = Convert.ToString(listShapesMesMgmtDevices.Count);
                //listShapesMesUplinkPorts[listShapesMesUplinkPorts.Count - 1].Data3 = Convert.ToString(listShapesMesMgmtDevices.Count);

                //////////////////////////  Group Ports & Device    //////////////////////////////////////////////////////////////////////////////////////

                //////////////////////////  Group LAN Chassis & Ports    ////////////////////////////

                /*
                intCurrentDeviceInGroup = 0;

                foreach (Visio.Shape objSingleLanDevice in listShapesLanDevices)
                {
                    vsoWindow.DeselectAll();
                    intCurrentDeviceInGroup++;
                    vsoWindow.Select(objSingleLanDevice, 2);
                    foreach (Visio.Shape objLanPort in listShapesLanPorts)
                    {
                        if (Convert.ToInt32(objLanPort.Data3) == intCurrentDeviceInGroup)
                        {
                            vsoWindow.Select(objLanPort, 2);
                        };

                    };
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };



                //////////////////////////  Group WAN Chassis & Ports    ////////////////////////////

                intCurrentDeviceInGroup = 0;

                foreach (Visio.Shape objSingleWanDevice in listShapesWanDevices)
                {
                    vsoWindow.DeselectAll();
                    intCurrentDeviceInGroup++;
                    vsoWindow.Select(objSingleWanDevice, 2);
                    foreach (Visio.Shape objWanPort in listShapesWanPorts)
                    {
                        if (Convert.ToInt32(objWanPort.Data3) == intCurrentDeviceInGroup)
                        {
                            vsoWindow.Select(objWanPort, 2);
                        };

                    };
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };



                //////////////////////////  Group 100G Bypass & Ports    ////////////////////////////

                intCurrentDeviceInGroup = 0;

                foreach (Visio.Shape objSingleBypass in listShapesBypass100Devices)
                {
                    vsoWindow.DeselectAll();
                    intCurrentDeviceInGroup++;
                    vsoWindow.Select(objSingleBypass, 2);
                    foreach (Visio.Shape objBypassNetPort in listShapesBypass100NetPorts)
                    {
                        //Console.WriteLine($"Device: {intCurrentDeviceInGroup}, Device_In_Data: {objBypassNetPort.Data1}, Port: {objBypassNetPort.Data2}");
                        if (Convert.ToInt32(objBypassNetPort.Data1) == intCurrentDeviceInGroup)
                        {
                            //  Console.Write($"Found; ");
                            vsoWindow.Select(objBypassNetPort, 2);
                            // Console.WriteLine($"Done");
                        };

                    };
                    foreach (Visio.Shape objBypassNetPort in listShapesBypass100MonPorts)
                    {
                        //Console.WriteLine($"Device: {intCurrentDeviceInGroup}, Device_In_Data: {objBypassNetPort.Data1}, Port: {objBypassNetPort.Data2}");
                        if (Convert.ToInt32(objBypassNetPort.Data1) == intCurrentDeviceInGroup)
                        {
                            //  Console.Write($"Found; ");
                            vsoWindow.Select(objBypassNetPort, 2);
                            // Console.WriteLine($"Done");
                        };

                    };
                    //Console.WriteLine("Check Selection 1");
                    vsoSelection = vsoWindow.Selection;
                    //Console.WriteLine("Check Selection 2");
                    vsoSelection.Group();
                    //Console.WriteLine("Check Selection 3");
                };

                //////////////////////////  Group IS40 Bypass & Ports    ////////////////////////////
                ///
                intCurrentDeviceInGroup = 0;

                foreach (Visio.Shape objSingleBypass in listShapesBypass10Devices)
                {
                    vsoWindow.DeselectAll();
                    intCurrentDeviceInGroup++;
                    vsoWindow.Select(objSingleBypass, 2);
                    foreach (Visio.Shape objBypassNetPort in listShapesBypass10NetPorts)
                    {
                        //Console.WriteLine($"Device: {intCurrentDeviceInGroup}, Device_In_Data: {objBypassNetPort.Data1}, Port: {objBypassNetPort.Data2}");
                        if (Convert.ToInt32(objBypassNetPort.Data1) == intCurrentDeviceInGroup)
                        {
                            //  Console.Write($"Found; ");
                            vsoWindow.Select(objBypassNetPort, 2);
                            // Console.WriteLine($"Done");
                        };

                    };
                    foreach (Visio.Shape objBypassNetPort in listShapesBypass10MonPorts)
                    {
                        //Console.WriteLine($"Device: {intCurrentDeviceInGroup}, Device_In_Data: {objBypassNetPort.Data1}, Port: {objBypassNetPort.Data2}");
                        if (Convert.ToInt32(objBypassNetPort.Data1) == intCurrentDeviceInGroup)
                        {
                            //  Console.Write($"Found; ");
                            vsoWindow.Select(objBypassNetPort, 2);
                            // Console.WriteLine($"Done");
                        };

                    };

                    //listDeviceMgmtPorts
                    foreach (Visio.Shape objMgmtPort in listDeviceMgmtPorts)
                    {
                        if (objMgmtPort.Data3 == objSingleBypass.Data3)
                        {
                            vsoWindow.Select(objMgmtPort, 2);
                        };

                    };


                    //listBypassHydraConnectors
                    foreach (Visio.Shape objSingleHydraConnector in listBypassHydraConnectors)
                    {
                        //Console.WriteLine($"Bypass {objSingleBypass.Data3}, Hydra {objSingleHydraConnector.Data3}");
                        if (objSingleHydraConnector.Data3 == objSingleBypass.Data3)
                        {
                            //Console.WriteLine("Check Data3");
                            vsoWindow.Select(objSingleHydraConnector, 2);
                        };

                    };

                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };

                //listShapesIbs1UpDevices


                //////////////////////////  Group IBS1UP Bypass & Ports    ////////////////////////////
                ///
                intCurrentDeviceInGroup = 0;

                foreach (Visio.Shape objSingleBypass in listShapesIbs1UpDevices)
                {
                    vsoWindow.DeselectAll();
                    intCurrentDeviceInGroup++;
                    vsoWindow.Select(objSingleBypass, 2);
                    foreach (Visio.Shape objBypassNetPort in listShapesIbs1upNetPorts)
                    {
                        //Console.WriteLine($"Device: {intCurrentDeviceInGroup}, Device_In_Data: {objBypassNetPort.Data1}, Port: {objBypassNetPort.Data2}");
                        if (Convert.ToInt32(objBypassNetPort.Data1) == intCurrentDeviceInGroup)
                        {
                            //  Console.Write($"Found; ");
                            vsoWindow.Select(objBypassNetPort, 2);
                            // Console.WriteLine($"Done");
                        };

                    };
                    foreach (Visio.Shape objBypassNetPort in listShapesIbs1upMonPorts)
                    {
                        //Console.WriteLine($"Device: {intCurrentDeviceInGroup}, Device_In_Data: {objBypassNetPort.Data1}, Port: {objBypassNetPort.Data2}");
                        if (Convert.ToInt32(objBypassNetPort.Data1) == intCurrentDeviceInGroup)
                        {
                            //  Console.Write($"Found; ");
                            vsoWindow.Select(objBypassNetPort, 2);
                            // Console.WriteLine($"Done");
                        };

                    };

                    //listDeviceMgmtPorts
                    foreach (Visio.Shape objMgmtPort in listDeviceMgmtPorts)
                    {
                        if (objMgmtPort.Data3 == objSingleBypass.Data3)
                        {
                            vsoWindow.Select(objMgmtPort, 2);
                        };

                    };


                    //listBypassHydraConnectors
                    foreach (Visio.Shape objSingleHydraConnector in listBypassHydraConnectors)
                    {
                        if (objSingleHydraConnector.Data3 == objSingleBypass.Data3)
                        {
                            vsoWindow.Select(objSingleHydraConnector, 2);
                        };

                    };

                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };


            

                //Zhopa

                //////////////////////////  Group 100G Balancers & Ports    ////////////////////////////


                //intCurrentDeviceInGroup = 0;


                foreach (Visio.Shape objSingleBalancer in listShapesBalancerDevices)
                {
                    vsoWindow.DeselectAll();
                    //intCurrentDeviceInGroup++;
                    vsoWindow.Select(objSingleBalancer, 2);
                    foreach (Visio.Shape objBalancerPort in listShapesBalancerDownlinkPorts)
                    {
                        if (objBalancerPort.Data3 == objSingleBalancer.Data3)
                        {
                            vsoWindow.Select(objBalancerPort, 2);
                        };

                    };
                    foreach (Visio.Shape objBalancerPort in listShapesBalancerUplinkPorts)
                    {
                        if (objBalancerPort.Data3 == objSingleBalancer.Data3)
                        {
                            vsoWindow.Select(objBalancerPort, 2);
                        };

                    };
        
                
                
                foreach (Visio.Shape objBalancerPort in listDeviceMgmtPorts)
                    {
                        //Console.WriteLine($"Balancer {intCurrentDeviceInGroup}, Mgmt_Port {objBalancerPort.Data3}");
                        if (objBalancerPort.Data3 == objSingleBalancer.Data3)
                        {
                            //  Console.WriteLine($"Check 1");
                            vsoWindow.Select(objBalancerPort, 2);
                            //  Console.WriteLine($"Check 2");
                        };

                    };

                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };



                //////////////////////////  Group 10G Filters & Ports    ////////////////////////////

                if (!boolNoBalancer)
                {

                    intCurrentDeviceInGroup = 0;


                    foreach (Visio.Shape objSingleFilter in listShapesFilter4160Devices)
                    {
                        vsoWindow.DeselectAll();
                        intCurrentDeviceInGroup++;
                        vsoWindow.Select(objSingleFilter, 2);
                        foreach (Visio.Shape objFilterPort in listShapesFilterPorts)
                        {
                            if (Convert.ToInt32(objFilterPort.Data1) == intCurrentDeviceInGroup)
                            {
                                vsoWindow.Select(objFilterPort, 2);
                            };

                        };



                        foreach (Visio.Shape objFilterHydra in listFilterHydraConnectors)
                        {
                            if (Convert.ToInt32(objFilterHydra.Data1) == intCurrentDeviceInGroup)
                            {
                                vsoWindow.Select(objFilterHydra, 2);
                            };

                        };

                        //listDeviceMgmtPorts[listDeviceMgmtPorts.Count - 1].Data3 = strCurrentDeviceHostname;
                        //listDeviceLogPorts.Add(page1.DrawRectangle);

                        foreach (Visio.Shape objMgmtPort in listDeviceMgmtPorts)
                        {
                            if (objMgmtPort.Data3 == objSingleFilter.Data3)
                            {
                                vsoWindow.Select(objMgmtPort, 2);
                            };
                        };

                        foreach (Visio.Shape objLogPort in listDeviceLogPorts)
                        {
                            if (objLogPort.Data3 == objSingleFilter.Data3)
                            {
                                vsoWindow.Select(objLogPort, 2);
                            };
                        };

                        vsoSelection = vsoWindow.Selection;
                        vsoSelection.Group();

                    };
                }

            //////////////////////////  Group Single Filter & Ports    ////////////////////////////



            /*

                        //////////////////////////  Template    /////////////////////////////////////



                        Visio.Selection vsoSelection;
                        Visio.Window vsoWindow;

                        vsoWindow = app.ActiveWindow;

                        int intCurrentDevice = 0;


                        foreach (Visio.Shape objSingleBypass in listShapesBypass100Devices)
                        {
                            vsoWindow.DeselectAll();
                            intCurrentDevice++;
                            vsoWindow.Select(objSingleBypass, 2);
                            foreach (Visio.Shape objBypassNetPort in listShapesBypass100NetPorts)
                            {
                                Console.WriteLine($"Device: {intCurrentDevice}, Device_In_Data: {objBypassNetPort.Data1}, Port: {objBypassNetPort.Data2}");
                                if (Convert.ToInt32(objBypassNetPort.Data1) == intCurrentDevice)
                                {
                                    Console.Write($"Found; ");
                                    vsoWindow.Select(objBypassNetPort, 2);
                                    Console.WriteLine($"Done");
                                };

                            };
                            vsoSelection = vsoWindow.Selection;
                            vsoSelection.Group();
                        };


                        */



                //doubStartPointNextShapeY -= 8.7;

                //  };
                /*

                /////////////////////////////////// Draw Connectors (only for test) //////////////////////////////////////////////



                foreach (Visio.Shape objLanPort in listShapesLanPorts)
                {
                    foreach (Visio.Shape objWanPort in listShapesWanPorts)
                    {
                        if (Convert.ToInt32(objLanPort.Data1) + 1 == Convert.ToInt32(objWanPort.Data1))
                            objLanPort.AutoConnect(objWanPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                    };
                };

                */




                //Add Rectangles
                //Visio.Shape rect1 = page1.DrawRectangle(8, 4, 10, 6);
                //Visio.Shape rect2 = page1.DrawRectangle(3, 1, 4, 2);
                //rect1.Text = @"Rect1";
                //rect2.Text = @"Rect2";

                //rect1.get_Cells("FillForegnd").FormulaU = "=RGB(150,10,0)";                      //Fill Cell with Color

                //rect1.CellsU["LineColor"].FormulaForceU = "THEMEGUARD(RGB(255,0,0))";           //Line Color
                //rect1.CellsU["RowFill"].FormulaForceU = "THEMEGUARD(RGB(255,0,0))";

                //Add Connector
                //rect1.AutoConnect(rect2, Visio.VisAutoConnectDir.visAutoConnectDirNone);


                //SaveAs
                //doc.SaveAs(@"c:\Users\v.patrukhachev\Documents\Scripts\Output.vsdx");





                doc.SaveAs(strVsdFilePath);


                //Quit
                doc.Close();
                app.Quit();


                /////////////////////////   Fill Cables ////////////////////////

                Excel.Range formatRange;
                Excel.Borders border;


                //Visio.Shape Connect1 = page1.DropConnected(arrShapesBypass100NetPorts[intCurrentOverallLinkNumber, 1], arrShapesWanPorts[intCurrentOverallLinkNumber], Visio.VisAutoConnectDir.visAutoConnectDirNone);


                /*
                formatRange = xlWorksheet21.get_Range("a1").EntireRow.EntireColumn;
                formatRange.NumberFormat = "@";
                formatRange = xlWorksheet22.get_Range("a1").EntireRow.EntireColumn;
                formatRange.NumberFormat = "@";
                formatRange = xlWorksheet23.get_Range("a1").EntireRow.EntireColumn;
                formatRange.NumberFormat = "@";
                formatRange = xlWorksheet24.get_Range("a1").EntireRow.EntireColumn;
                formatRange.NumberFormat = "@";
                formatRange = xlWorksheet25.get_Range("a1").EntireRow.EntireColumn;
                formatRange.NumberFormat = "@";
                */
                formatRange = xlWorksheet31.get_Range("a1").EntireRow.EntireColumn;
                formatRange.NumberFormat = "@";

                /*
                intJournalCurrentRow = 0;

                //Console.WriteLine("Journal 1");


                foreach (Dictionary<string, string> dictCableRecord in listCableJournal_1)
                {
                    intJournalCurrentRow++;
                    xlWorksheet21.Cells[intJournalCurrentRow + 2, 1] = intJournalCurrentRow;
                    xlWorksheet21.Cells[intJournalCurrentRow + 2, 2] = dictCableRecord["Device_A_Name"];
                    xlWorksheet21.Cells[intJournalCurrentRow + 2, 3] = dictCableRecord["Port_A_Name"];
                    xlWorksheet21.Cells[intJournalCurrentRow + 2, 5] = dictCableRecord["Device_B_Name"];
                    xlWorksheet21.Cells[intJournalCurrentRow + 2, 6] = dictCableRecord["Port_B_Name"];
                    //Console.WriteLine($"{dictCableRecord["Device_A_Name"]}   {dictCableRecord["Port_A_Name"]}   {dictCableRecord["Device_B_Name"]}   {dictCableRecord["Port_B_Name"]}");
                };


                //Console.WriteLine("Journal 2");

                intJournalCurrentRow = 0;

                foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_Bypass_Filter)
                {
                    //Console.WriteLine($"{dictCableRecord.Keys}     {dictCableRecord.Values}");
                    //if (dictCableRecord.ContainsKey("Device_B_Name"))
                    //{
                    // Console.WriteLine("Check1");
                    //Console.WriteLine($"{dictCableRecord["Device_A_Name"]}   {dictCableRecord["Port_A_Name"]}   {dictCableRecord["Device_B_Name"]}   {dictCableRecord["Port_B_Name"]}");
                    //Console.WriteLine($"{dictCableRecord["Device_A_Name"]}   {dictCableRecord["Port_A_Name"]}");
                    intJournalCurrentRow++;
                    xlWorksheet22.Cells[intJournalCurrentRow + 2, 1] = intJournalCurrentRow;
                    xlWorksheet22.Cells[intJournalCurrentRow + 2, 2] = dictCableRecord["Device_A_Name"];
                    xlWorksheet22.Cells[intJournalCurrentRow + 2, 3] = dictCableRecord["Port_A_Name"];
                    xlWorksheet22.Cells[intJournalCurrentRow + 2, 5] = dictCableRecord["Device_B_Name"];
                    xlWorksheet22.Cells[intJournalCurrentRow + 2, 6] = dictCableRecord["Port_B_Name"];
                    //Console.WriteLine("Check2");
                    //Console.WriteLine($"{dictCableRecord["Device_A_Name"]}   {dictCableRecord["Port_A_Name"]}   {dictCableRecord["Device_B_Name"]}   {dictCableRecord["Port_B_Name"]}"); 
                    //};

                };

                //Console.WriteLine("Journal 3");

                intJournalCurrentRow = 0;

                if (!boolNoBalancer)
                {
                    //Console.WriteLine($"CJ3 Records: {listCableJournal_3.Count}");
                    foreach (Dictionary<string, string> dictCableRecord in listCableJournal_3)
                    {
                        if (dictCableRecord.ContainsKey("Device_B_Name"))
                        {
                            intJournalCurrentRow++;
                            xlWorksheet23.Cells[intJournalCurrentRow + 2, 1] = intJournalCurrentRow;
                            xlWorksheet23.Cells[intJournalCurrentRow + 2, 2] = dictCableRecord["Device_A_Name"];
                            xlWorksheet23.Cells[intJournalCurrentRow + 2, 3] = dictCableRecord["Port_A_Name"];
                            //Console.WriteLine($"Device_A: {dictCableRecord["Device_A_Name"]}, Port_A: {dictCableRecord["Port_A_Name"]}, Device_B: {dictCableRecord["Device_B_Name"]}, Port_B: {dictCableRecord["Port_B_Name"]}");
                            xlWorksheet23.Cells[intJournalCurrentRow + 2, 5] = dictCableRecord["Device_B_Name"];
                            xlWorksheet23.Cells[intJournalCurrentRow + 2, 6] = dictCableRecord["Port_B_Name"];
                            // Console.WriteLine($"{dictCableRecord["Device_A_Name"]}  {dictCableRecord["Port_A_Name"]}  {dictCableRecord["Device_B_Name"]}  {dictCableRecord["Port_B_Name"]}");
                        };

                    };
                }

                // Console.WriteLine("Journal 4");

                intJournalCurrentRow = 0;

                foreach (Dictionary<string, string> dictCableRecord in listCableJournal_4)
                {
                    //Console.WriteLine($"{dictCableRecord["Device_A_Name"]}   {dictCableRecord["Port_A_Name"]}   {dictCableRecord["Device_B_Name"]}   {dictCableRecord["Port_B_Name"]}");
                    // Console.WriteLine($"{dictCableRecord["Device_A_Name"]}   {dictCableRecord["Port_A_Name"]}");
                    intJournalCurrentRow++;
                    xlWorksheet24.Cells[intJournalCurrentRow + 3, 1] = intJournalCurrentRow;
                    xlWorksheet24.Cells[intJournalCurrentRow + 3, 2] = dictCableRecord["Operator_LAN_Device"];
                    xlWorksheet24.Cells[intJournalCurrentRow + 3, 3] = dictCableRecord["Operator_LAN_Port"];
                    xlWorksheet24.Cells[intJournalCurrentRow + 3, 4] = dictCableRecord["Bypass_Chassis"];
                    xlWorksheet24.Cells[intJournalCurrentRow + 3, 5] = dictCableRecord["Bypass_LAN_Port"];
                    xlWorksheet24.Cells[intJournalCurrentRow + 3, 6] = dictCableRecord["Bypass_WAN_Port"];
                    xlWorksheet24.Cells[intJournalCurrentRow + 3, 7] = dictCableRecord["Operator_WAN_Device"];
                    xlWorksheet24.Cells[intJournalCurrentRow + 3, 8] = dictCableRecord["Operator_WAN_Port"];
                    //Console.WriteLine("Check2");
                    //Console.WriteLine($"{dictCableRecord["Device_A_Name"]}   {dictCableRecord["Port_A_Name"]}   {dictCableRecord["Device_B_Name"]}   {dictCableRecord["Port_B_Name"]}"); 
                    //};

                };


                // Console.WriteLine("Journal 5");

                intJournalCurrentRow = 0;

                foreach (Dictionary<string, string> dictCableRecord in listCableJournal_Management)
                {
                    //Console.WriteLine($"{dictCableRecord["Device_A"]}   {dictCableRecord["Port_A"]}   {dictCableRecord["Device_B"]}   {dictCableRecord["Port_B"]}");
                    //Console.WriteLine($"{dictCableRecord["Device_A"]}   {dictCableRecord["Port_A"]}");
                    intJournalCurrentRow++;
                    xlWorksheet25.Cells[intJournalCurrentRow + 2, 1] = intJournalCurrentRow;
                    xlWorksheet25.Cells[intJournalCurrentRow + 2, 2] = dictCableRecord["Device_A"];
                    xlWorksheet25.Cells[intJournalCurrentRow + 2, 3] = dictCableRecord["Port_A"];
                    xlWorksheet25.Cells[intJournalCurrentRow + 2, 5] = dictCableRecord["Device_B"];
                    xlWorksheet25.Cells[intJournalCurrentRow + 2, 6] = dictCableRecord["Port_B"];
                    //Console.WriteLine("Check");
                    //Console.WriteLine($"{dictCableRecord["Device_A"]}   {dictCableRecord["Port_A"]}   {dictCableRecord["Device_B"]}   {dictCableRecord["Port_B"]}"); 
                    //};

                };

                */







                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ !!!! Большой КЖ !!!!! ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Console.WriteLine("~~~~~~~~~~~~~~   Journal Common  ~~~~~~~~~~~~~~~~");
                // QWERTY
                Console.WriteLine("~~~  LAN-Bypass  ~~~");

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




                Console.WriteLine("~~~  WAN-Bypass  ~~~");

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




                Console.WriteLine("~~~  Bypass-Balancer  ~~~");

                intJournalCurrentRow++;

                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);

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
                    };
                }

                else
                {
                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Байпасы - Балансировщики";
                    foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_Bypass_Balancer)
                    {
                        intJournalCurrentRow++;
                        xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"]; ;
                        xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Device_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Port_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Device_B_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Port_B_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 7] = "FT-QSFP+/4SFP+CabA-2";
                    };

                    Console.WriteLine("~~~  Balancer-Filter  ~~~");
                    intJournalCurrentRow++;
                    xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                    xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);

                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Балансировщики - Фильтры";

                    foreach (Dictionary<string, string> dictCableRecord in list_CableJournal_Balancer_Filter)
                    //foreach (Dictionary<string, string> dictCableRecord in listCableJournal_3)
                    {
                        intJournalCurrentRow++;
                        xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"]; ;
                        xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Device_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Port_A_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Device_B_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Port_B_Name"];
                        xlWorksheet31.Cells[intJournalCurrentRow, 7] = "FT-QSFP+/4SFP+CabA-2";
                    };
                };





                Console.WriteLine("~~~  Служебное  ~~~");

                intJournalCurrentRow++;

                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 11]].Merge();
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorksheet31.Range[xlWorksheet31.Cells[intJournalCurrentRow, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Bisque);
                xlWorksheet31.Cells[intJournalCurrentRow, 1] = "Сегмент управления и логирования";


                //Console.WriteLine("~~~  Management & Log  ~~~");

                foreach (Dictionary<string, string> dictCableRecord in listCableJournal_Management)
                {
                    intJournalCurrentRow++;
                    xlWorksheet31.Cells[intJournalCurrentRow, 1] = dictCableRecord["Cable_Number"]; ;
                    xlWorksheet31.Cells[intJournalCurrentRow, 2] = dictCableRecord["Cable_Name"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 3] = dictCableRecord["Device_A"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 4] = dictCableRecord["Port_A"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 5] = dictCableRecord["Device_B"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 6] = dictCableRecord["Port_B"];
                    xlWorksheet31.Cells[intJournalCurrentRow, 7] = dictCableRecord["Cable_Type"];
                };






                xlWorksheet31.Range[xlWorksheet31.Cells[2, 1], xlWorksheet31.Cells[intJournalCurrentRow, 1]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

                /*

                   */



                Console.WriteLine("!!!!!!!!!!!!!! End !!!!!!!!!!!!!!!!");












                /*

                formatRange = xlWorksheet21.UsedRange;
                border = formatRange.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 3d;
                formatRange.Columns.AutoFit();

                formatRange = xlWorksheet22.UsedRange;
                border = formatRange.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 3d;
                formatRange.Columns.AutoFit();

                formatRange = xlWorksheet23.UsedRange;
                border = formatRange.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 3d;
                formatRange.Columns.AutoFit();

                formatRange = xlWorksheet24.UsedRange;
                border = formatRange.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 3d;
                formatRange.Columns.AutoFit();

                formatRange = xlWorksheet25.UsedRange;
                border = formatRange.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 3d;
                formatRange.Columns.AutoFit();

                */


                formatRange = xlWorksheet31.UsedRange;
                border = formatRange.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 3d;
                formatRange.Columns.AutoFit();


                xlApp2.DisplayAlerts = false;
                xlWorkbook2.SaveAs(strExcelCablesFilePath);
                xlWorkbook2.Close();
                xlApp2.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkbook2);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp2);

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////



                //Console Output
                //Console.WriteLine("Check");

                //Console.WriteLine(strXlsxFilePath);
            }
            catch (Exception err)
            {
                Console.WriteLine("Error: {0}", err.Message);
            }
            finally
            {
                Console.Write("\nPress any key to continue ...");
                Console.ReadKey();
            }
        }


    }


}





/*
*/
