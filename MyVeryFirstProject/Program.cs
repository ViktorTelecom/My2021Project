using System;
using Visio = Microsoft.Office.Interop.Visio;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
//using System.Security.Permissions;
using System.Threading;
using System.Collections.Generic;
//using IVisio.Window;


namespace VeryFirstProject
{

    class Program
    {


        static void Main(string[] args)
        {

            //Open File Fialog
            string strSelectedPath = "";

            Thread thread1 = new Thread((ThreadStart)(() => {
                OpenFileDialog newOpenFileDialog1 = new OpenFileDialog();

                newOpenFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
                newOpenFileDialog1.FilterIndex = 2;
                newOpenFileDialog1.RestoreDirectory = true;

                if (newOpenFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    strSelectedPath = newOpenFileDialog1.FileName;
                }
            }));


            // Run your code from a thread that joins the STA Thread



            string strVsdFilePath;


            try
            {

                thread1.SetApartmentState(ApartmentState.STA);
                thread1.Start();
                thread1.Join();

                //Console.WriteLine("Check");

                Console.WriteLine("File Path: {0}", strSelectedPath);


                strVsdFilePath = strSelectedPath.Replace("xlsx", "vsdx");



                Console.WriteLine($"FileName: {strVsdFilePath}");

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                //open excel
                Excel.Application xlApp = new Excel.Application();
                //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"c:\DCOA\excel_template.xlsx");
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(strSelectedPath);
                Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int intTotalRows = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                //Console.WriteLine("Rows: {0}", rowCount);
                Console.WriteLine($"Rows: {intTotalRows}");











                //read cells    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                List<Dictionary<string, object>> listLanDevices = new List<Dictionary<string, object>>();
                List<Dictionary<string, object>> listWanDevices = new List<Dictionary<string, object>>();
                List<Dictionary<string, object>> listLanPorts = new List<Dictionary<string, object>>();
                List<Dictionary<string, object>> listWanPorts = new List<Dictionary<string, object>>();
                List<Dictionary<string, int>> listPortPairs = new List<Dictionary<string, int>>();

                List<Visio.Shape> listShapesLanDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesWanDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesLanPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesWanPorts = new List<Visio.Shape>();
                List<Visio.Shape> listConnectorLines = new List<Visio.Shape>();

                int intCurrentLanPortsCounter;
                int intCurrentWanPortsCounter;
                string strCurrentLanHostname;
                string strCurrentWanHostname;
                string strCurrentLanPortName;
                string strCurrentWanPortName;
                string strCurrentLinkType;

                bool boolLanObjectNotFound;
                bool boolWanObjectNotFound;

                int intCurrentGlobalDeviceIndex;
                int intCurrentGlobalPortIndex;
                int intCurrentGlobalDeviceIndexForPort;
                int intCurrentDeviceInGroup;

                intCurrentGlobalDeviceIndex = 0;
                intCurrentGlobalPortIndex = 0;
                intCurrentGlobalDeviceIndexForPort = 0;


                int intLinkCounter100 = 0;
                int intLinkCounter40 = 0;
                int intLinkCounter10 = 0;
                int intLinkCounter1Fiber = 0;
                int intLinkCounter1Copper = 0;

                double doubSummaryBandwidth;

                int intTotalFilters;


                double doubBandwidthOnFilter4160 = 71;


                doubSummaryBandwidth = Convert.ToDouble(((Excel.Range)xlWorksheet.Cells[2, 7]).Value2.ToString());          //Read BW and calculate filters quantity
                intTotalFilters = Convert.ToInt32(doubSummaryBandwidth/ doubBandwidthOnFilter4160);
                if ((doubSummaryBandwidth % doubBandwidthOnFilter4160 > 0) & (doubSummaryBandwidth > doubBandwidthOnFilter4160)) intTotalFilters++;
                if (doubSummaryBandwidth < doubBandwidthOnFilter4160) intTotalFilters = 1;



                Console.WriteLine($"Total Bandwidth: {doubSummaryBandwidth}, Filters: {intTotalFilters}");

                int intLocalPortCounter = 0;

                for (int intCurrentRow = 2; intCurrentRow <= intTotalRows; intCurrentRow++)
                {
                    boolLanObjectNotFound = true;
                    boolWanObjectNotFound = true;
                    strCurrentLanHostname = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 1]).Value2.ToString();
                    strCurrentLanPortName = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 2]).Value2.ToString();
                    strCurrentWanHostname = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 3]).Value2.ToString();
                    strCurrentWanPortName = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 4]).Value2.ToString();
                    strCurrentLinkType = ((Excel.Range)xlWorksheet.Cells[intCurrentRow, 5]).Value2.ToString();



                    //////////////////////////////////////////// Count Ports of Each Type ////////////////////////////////////////////////

                    switch (strCurrentLinkType)
                    {
                        case "100":
                            intLinkCounter100++;
                            break;
                        case "40":
                            intLinkCounter40++;
                            break;
                        case "10":
                            intLinkCounter10++;
                            break;
                        case "1Fiber":
                            intLinkCounter1Fiber++;
                            break;
                        case "1Copper":
                            intLinkCounter1Copper++;
                            break;
                    }





                    //////////////////////////////////////////// LAN Devices ////////////////////////////////////////////////

                    // Console.WriteLine();
                    // Console.WriteLine("##############################################################");
                    // Console.WriteLine($"Row {intCurrentRow}");
                    // Console.WriteLine("##############################################################");

                    if (listLanDevices.Count > 0)
                    {
                        foreach (Dictionary<string, object> dictLanDevices in listLanDevices)
                        {
                            if (dictLanDevices.ContainsValue(strCurrentLanHostname))
                            {
                                intCurrentLanPortsCounter = Convert.ToInt32(dictLanDevices["Ports_Number"]);
                                intCurrentLanPortsCounter++;
                                dictLanDevices["Ports_Number"] = intCurrentLanPortsCounter;
                                //    Console.WriteLine($"Counter LAN: {intCurrentLanPortsCounter}");
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
                        // Console.WriteLine($"~~~~~~~~~~~~~~~~~~~~~~");
                        listLanDevices[listLanDevices.Count - 1].Add("Device_Name", (strCurrentLanHostname));
                        listLanDevices[listLanDevices.Count - 1].Add("Ports_Number", (1));
                        listLanDevices[listLanDevices.Count - 1].Add("Device_Index", intCurrentGlobalDeviceIndex);
                        intCurrentGlobalDeviceIndexForPort = intCurrentGlobalDeviceIndex;

                        //    Console.WriteLine($"Counter LAN: 1");
                        //    Console.WriteLine($"{strCurrentLanHostname} - Added Record: {listLanDevices.Count - 1}.");
                        //    Console.WriteLine();
                    };

                    ///////////////////////////////////// LAN Ports   ////////////////////////////////////////////////////

                    listLanPorts.Add(new Dictionary<string, object>());
                    intCurrentGlobalPortIndex++;

                    switch (strCurrentLinkType)
                    {
                        case "100":
                            intLocalPortCounter = 2 * intLinkCounter100 - 1;
                            break;
                        case "40":
                            intLocalPortCounter = 2 * intLinkCounter40 - 1;
                            break;
                        case "10":
                            intLocalPortCounter = 2 * intLinkCounter10 - 1;
                            break;
                        case "1Fiber":
                            intLocalPortCounter = 2 * intLinkCounter1Fiber - 1;
                            break;
                        case "1Copper":
                            intLocalPortCounter = 2 * intLinkCounter1Copper - 1;
                            break;
                    }

                    //Console.WriteLine($"~~~~~~~~~~~~~~~~~~~~~~");
                    listLanPorts[listLanPorts.Count - 1].Add("Device_Index", (intCurrentGlobalDeviceIndexForPort));
                    listLanPorts[listLanPorts.Count - 1].Add("Port_Name", (strCurrentLanPortName));
                    listLanPorts[listLanPorts.Count - 1].Add("Port_Index", (intLocalPortCounter));
                    listLanPorts[listLanPorts.Count - 1].Add("Link_Type", strCurrentLinkType);







                    /*                              Шаблон для Case Link Type
                    switch (value)
                    {
                        case 1:
                        case 2:
                        case 3:
                            // Do Something
                            break;
                        case 4:
                        case 5:
                        case 6:
                            // Do Something
                            break;
                        default:
                            // Do Something
                            break;
                    }


                    case 1: case 2: case 3:
                    // Do something
                    break;

                    */





                    //////////////////////////////////////// WAN Devices ////////////////////////////////////////////////


                    if (listWanDevices.Count > 0)
                    {
                        foreach (Dictionary<string, object> dictWanDevices in listWanDevices)
                        {
                            if (dictWanDevices.ContainsValue(strCurrentWanHostname))
                            {
                                intCurrentWanPortsCounter = Convert.ToInt32(dictWanDevices["Ports_Number"]);
                                intCurrentWanPortsCounter++;
                                dictWanDevices["Ports_Number"] = intCurrentWanPortsCounter;
                                //    Console.WriteLine($"Counter WAN: {intCurrentWanPortsCounter}");
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


                        // Console.WriteLine($"~~~~~~~~~~~~~~~~~~~~~~");
                        listWanDevices[listWanDevices.Count - 1].Add("Device_Name", (strCurrentWanHostname));
                        listWanDevices[listWanDevices.Count - 1].Add("Ports_Number", (1));
                        listWanDevices[listWanDevices.Count - 1].Add("Device_Index", intCurrentGlobalDeviceIndex);
                        intCurrentGlobalDeviceIndexForPort = intCurrentGlobalDeviceIndex;

                        //Console.WriteLine($"Counter WAN: 1");
                        //Console.WriteLine($"{strCurrentWanHostname} - Added Record: {listWanDevices.Count - 1}.");
                        //Console.WriteLine();
                    };




                    /////////////////////////////// WAN Ports   ////////////////////////////////////////////////////

                    listWanPorts.Add(new Dictionary<string, object>());
                    intCurrentGlobalPortIndex++;

                    switch (strCurrentLinkType)
                    {
                        case "100":
                            intLocalPortCounter = 2 * intLinkCounter100;
                            break;
                        case "40":
                            intLocalPortCounter = 2 * intLinkCounter40;
                            break;
                        case "10":
                            intLocalPortCounter = 2 * intLinkCounter10;
                            break;
                        case "1Fiber":
                            intLocalPortCounter = 2 * intLinkCounter1Fiber;
                            break;
                        case "1Copper":
                            intLocalPortCounter = 2 * intLinkCounter1Copper;
                            break;
                    }


                    //Console.WriteLine($"~~~~~~~~~~~~~~~~~~~~~~");
                    listWanPorts[listWanPorts.Count - 1].Add("Device_Index", (intCurrentGlobalDeviceIndexForPort));
                    listWanPorts[listWanPorts.Count - 1].Add("Port_Name", (strCurrentWanPortName));
                    listWanPorts[listWanPorts.Count - 1].Add("Port_Index", (intLocalPortCounter));
                    listWanPorts[listWanPorts.Count - 1].Add("Link_Type", strCurrentLinkType);


                    /*

                    ///////////////////////////////  LAN-WAN Port Pairs (Unused)   ////////////////////////////////////////////////////

                    listPortPairs.Add(new Dictionary<string, int>());
                    listPortPairs[listLanPorts.Count - 1].Add("Index_Port_LAN", (intCurrentGlobalPortIndex-1));
                    listPortPairs[listLanPorts.Count - 1].Add("Index_Port_WAN", (intCurrentGlobalPortIndex));

                    */


                    //////////////  End Excel Pass  ///////////////////////////////

                };

                //close excel
                xlWorkbook.Close();
                xlApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);




                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                Console.WriteLine($"100G Ports: {intLinkCounter100}");
                Console.WriteLine($"40G Ports: {intLinkCounter40}");
                Console.WriteLine($"10G Ports: {intLinkCounter10}");
                Console.WriteLine($"1G Fiber Ports: {intLinkCounter1Fiber}");
                Console.WriteLine($"1G Copper Ports: {intLinkCounter1Copper}");



                int intTotalIs100Bypasses;
                intTotalIs100Bypasses = (intLinkCounter100) / 2;               //Calculate quantity of 100G (IS100) Bypasses  
                if ((intLinkCounter100) % 2 > 0) intTotalIs100Bypasses++;
                Console.WriteLine($"IS100 Number: {intTotalIs100Bypasses}");


                int intTotalIs40Bypasses;
                intTotalIs40Bypasses = (intLinkCounter40) / 3;               //Calculate quantity of 40G (IS100) Bypasses  
                if ((intLinkCounter40) % 3 > 0) intTotalIs40Bypasses++;
                // Console.WriteLine($"IS100 Number: {intTotalIs40Bypasses}");

                int intTotalIs10Bypasses;
                intTotalIs10Bypasses = (intLinkCounter10) / 6;
                if ((intLinkCounter10) % 6 > 0) intTotalIs10Bypasses++;

                Console.WriteLine($"IS40 Number: {intTotalIs40Bypasses} + {intTotalIs10Bypasses} = {intTotalIs40Bypasses + intTotalIs10Bypasses}");


                int intTotalIs1FiberBypasses;
                intTotalIs1FiberBypasses = (intLinkCounter1Fiber) / 4;               //Calculate quantity of 1G Fiber (IBS1U) Bypasses  
                if ((intLinkCounter1Fiber) % 4 > 0) intTotalIs1FiberBypasses++;


                int intTotalIs1CopperBypasses;
                intTotalIs1CopperBypasses = (intLinkCounter1Copper) / 4;               //Calculate quantity of 1G Copper (IBS1U) Bypasses  
                if ((intLinkCounter1Copper) % 4 > 0) intTotalIs1CopperBypasses++;

                Console.WriteLine($"IBS1U Number: {intTotalIs1FiberBypasses} + {intTotalIs1CopperBypasses} = {intTotalIs1FiberBypasses + intTotalIs1CopperBypasses}");

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                double doubStartPointNextShapeX;
                double doubStartPointNextShapeY;


                double doubDeviceStartPointX = 0;
                double doubDeviceStartPointY;
                double doubDeviceEndPointX;
                double doubDeviceEndPointY;

                int intPortsOnDevice;

                ///////////////////////////////////////////////////////////////////////////////////


                //Console.WriteLine($"~~~~~~~~~~~~~~~~~~~    LAN Device Dictionaries  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");

                doubStartPointNextShapeX = 1;
                doubStartPointNextShapeY = 1;

                int intListCurrentIndex;

                intListCurrentIndex = 0;

                foreach (Dictionary<string, object> dictLanDevices in listLanDevices)
                {
                    intPortsOnDevice = Convert.ToInt32(dictLanDevices["Ports_Number"]);
                    doubDeviceStartPointX = doubStartPointNextShapeX;
                    doubDeviceStartPointY = doubStartPointNextShapeY;
                    doubDeviceEndPointX = doubDeviceStartPointX + 0.25 * intPortsOnDevice;
                    doubDeviceEndPointY = doubDeviceStartPointY + 1;
                    doubStartPointNextShapeX = doubDeviceEndPointX + 1.5;
                    doubStartPointNextShapeY = doubDeviceStartPointY;

                    listLanDevices[intListCurrentIndex].Add("StartX", doubDeviceStartPointX);
                    listLanDevices[intListCurrentIndex].Add("StartY", doubDeviceStartPointY);
                    listLanDevices[intListCurrentIndex].Add("EndX", doubDeviceEndPointX);
                    listLanDevices[intListCurrentIndex].Add("EndY", doubDeviceEndPointY);


                    //Console.WriteLine("---");
                    foreach (KeyValuePair<string, object> kvp in dictLanDevices)
                    {

                        //  Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}.");
                    };
                    intListCurrentIndex++;
                };

                //Console.WriteLine();





                // Console.WriteLine($"~~~~~~~~~~~~~~~~~~~    WAN Device Dictionaries  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");

                doubStartPointNextShapeX = 1;
                doubStartPointNextShapeY = intTotalIs100Bypasses * 2.7 + intTotalIs40Bypasses * 4.5 + intTotalIs10Bypasses * 9 + intTotalIs1FiberBypasses * 6 + intTotalIs1CopperBypasses * 6;

                intListCurrentIndex = 0;

                foreach (Dictionary<string, object> dictWanDevices in listWanDevices)
                {

                    intPortsOnDevice = Convert.ToInt32(dictWanDevices["Ports_Number"]);
                    doubDeviceStartPointX = doubStartPointNextShapeX;
                    doubDeviceStartPointY = doubStartPointNextShapeY;
                    doubDeviceEndPointX = doubDeviceStartPointX + 0.2 * intPortsOnDevice;
                    doubDeviceEndPointY = doubDeviceStartPointY + 1;
                    doubStartPointNextShapeX = doubDeviceEndPointX + 1.5;
                    doubStartPointNextShapeY = doubDeviceStartPointY;

                    listWanDevices[intListCurrentIndex].Add("StartX", doubDeviceStartPointX);
                    listWanDevices[intListCurrentIndex].Add("StartY", doubDeviceStartPointY);
                    listWanDevices[intListCurrentIndex].Add("EndX", doubDeviceEndPointX);
                    listWanDevices[intListCurrentIndex].Add("EndY", doubDeviceEndPointY);



                    // Console.WriteLine("---");
                    foreach (KeyValuePair<string, object> kvp in dictWanDevices)
                    {




                        //    Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}.");
                    };
                    intListCurrentIndex++;
                };

                //  Console.WriteLine();


                //   Console.WriteLine($"~~~~~~~~~~~~~~~~~~~    LAN Port Dictionaries  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                foreach (Dictionary<string, object> dictLanPorts in listLanPorts)
                {
                    //   Console.WriteLine("---");
                    foreach (KeyValuePair<string, object> kvp in dictLanPorts)
                    {
                        //  Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}.");
                    };
                };

                // Console.WriteLine();




                // Console.WriteLine($"~~~~~~~~~~~~~~~~~~~    WAN Port Dictionaries  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                foreach (Dictionary<string, object> dictWanPorts in listWanPorts)
                {
                    //  Console.WriteLine("---");
                    foreach (KeyValuePair<string, object> kvp in dictWanPorts)
                    {
                        //    Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}.");
                    };
                };

                //  Console.WriteLine();


                //  Console.WriteLine($"~~~~~~~~~~~~~~~~~~~    LAN-WAN Port Pairs Dictionaries  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                foreach (Dictionary<string, int> dicPortPairs in listPortPairs)
                {
                    //  Console.WriteLine("---");
                    foreach (KeyValuePair<string, int> kvp in dicPortPairs)
                    {
                        //    Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}.");
                    };
                };

                //   Console.WriteLine();

                ////////////////////////////////    Math Calculations   //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



                bool boolNoBalancer = false;


                if (intLinkCounter10 > 0 & intLinkCounter10 <= 8) boolNoBalancer = true;

                //Console.WriteLine(boolNoBalancer);

                int intBalancersQuantity;


                intBalancersQuantity = (2 * intLinkCounter100 + 2 * intLinkCounter40 + intLinkCounter10 / 2) / 16;
                if ((2 * intLinkCounter100 + 2 * intLinkCounter40 + intLinkCounter10 / 2) % 16 > 0) intBalancersQuantity++;

                //Console.WriteLine($"Balancers: {intBalancersQuantity}");




                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////










                // Start Visio
                Visio.Application app = new Visio.Application();

                // Create a new document.
                Visio.Document doc = app.Documents.Add("");

                // The new document will have one page,
                // get the a reference to it.
                Visio.Page page1 = doc.Pages[1];

                // Add a second page.
                // Visio.Page page2 = doc.Pages.Add();

                // Name the pages. This is what is shown in the page tabs.
                page1.Name = "Test C# Drawing";
                //page2.Name = "Page-2";



                Visio.Selection vsoSelection;
                Visio.Window vsoWindow;

                vsoWindow = app.ActiveWindow;
















                /////////////////////////////////// Draw Rectangles //////////////////////////////////////////////

                //double doubPortStartPointX;
                //double doubPortStartPointY;
                //double doubPortEndPointX;
                //double doubPortEndPointY;

                double doubNextPortStartPointX;
                double doubNextPortStartPointY;

                double doubLanLastX;
                double doubWanLastX;

                string strShapeName;
                string strPortName;


                int intCurrentLanChassis = 0;

                int intCurrentWanChassis = 0;



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
                        // Console.WriteLine($"Device Index: {dictLanDevices["Device_Index"]}, Port Index: {dictLanPorts["Device_Index"]}.");
                        if (Convert.ToInt32(dictLanDevices["Device_Index"]) == Convert.ToInt32(dictLanPorts["Device_Index"]))
                        {
                            //Console.WriteLine("Port on Device.");
                            //strPortName = Convert.ToString(dictLanPorts["Port_Name"]) + " (" +  Convert.ToString(dictLanPorts["Port_Index"] + ")");
                            strPortName = Convert.ToString(dictLanPorts["Port_Name"]);
                            listShapesLanPorts.Add(page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.6, doubNextPortStartPointY + 0.2));
                            listShapesLanPorts[listShapesLanPorts.Count - 1].Data1 = Convert.ToString(dictLanPorts["Port_Index"]);
                            listShapesLanPorts[listShapesLanPorts.Count - 1].Data2 = Convert.ToString(dictLanPorts["Link_Type"]);
                            listShapesLanPorts[listShapesLanPorts.Count - 1].Data3 = Convert.ToString(intCurrentLanChassis);

                            listShapesLanPorts[listShapesLanPorts.Count - 1].Text = strPortName;
                            doubNextPortStartPointX += 0.2;
                            listShapesLanPorts[listShapesLanPorts.Count - 1].Rotate90();
                            listShapesLanPorts[listShapesLanPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        };

                    };


                    doubLanLastX = doubDeviceEndPointX;
                    //intListCurrentIndex++;
                };

                

                //intCurrentDeviceInGroup



                /////////////////////////////////////   Define Single Filter Type   //////////////////////////////////////
                ///

                string strSingleFilterType = "";

                int intPortsNumberOnSingleFilter = 0;


                switch (listShapesLanPorts.Count)
                {
                    case int n when (n > 6 && n <= 8):
                        strSingleFilterType = "4160";
                        intPortsNumberOnSingleFilter = 4;
                        Console.WriteLine($"I am 4160. {listShapesLanPorts.Count} links.");
                        break;

                    case int n when(n <= 6 && n >= 5 ):
                        strSingleFilterType = "4120";
                        intPortsNumberOnSingleFilter = 6;
                        Console.WriteLine($"I am 4120. {listShapesLanPorts.Count} links.");
                        break;

                    case int n when(n <= 4):
                        strSingleFilterType = "4080";
                        intPortsNumberOnSingleFilter = 8;
                        Console.WriteLine($"I am 4080. {listShapesLanPorts.Count} links.");
                        break;
                }











                /////////////////////////////////// Draw WAN Chassis //////////////////////////////////////////////

                double doubOldDevicesEndPointX;

                foreach (Dictionary<string, object> dictWanDevices in listWanDevices)
                {
                    intCurrentWanChassis++;
                    doubDeviceStartPointX = Convert.ToDouble(dictWanDevices["StartX"]) + 0.3;
                    doubDeviceStartPointY = Convert.ToDouble(dictWanDevices["StartY"]);
                    doubDeviceEndPointX = Convert.ToDouble(dictWanDevices["EndX"]) + 0.5;
                    doubDeviceEndPointY = Convert.ToDouble(dictWanDevices["EndY"]);

                    strShapeName = Convert.ToString(dictWanDevices["Device_Name"]);

                    listShapesWanDevices.Add(page1.DrawRectangle(doubDeviceStartPointX, doubDeviceStartPointY, doubDeviceEndPointX, doubDeviceEndPointY));
                    listShapesWanDevices[listShapesWanDevices.Count - 1].Text = strShapeName;
                    listShapesWanDevices[listShapesWanDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(169,169,169)";

                    doubOldDevicesEndPointX = doubDeviceEndPointX;

                    /////////////////////////////////// Draw WAN Ports //////////////////////////////////////////////

                    doubNextPortStartPointX = doubDeviceStartPointX - 0.1;
                    doubNextPortStartPointY = doubDeviceStartPointY - 0.2;


                    foreach (Dictionary<string, object> dictWanPorts in listWanPorts)
                    {
                        //  Console.WriteLine($"Device Index: {dictWanDevices["Device_Index"]}, Port Index: {dictWanPorts["Device_Index"]}.");
                        if (Convert.ToInt32(dictWanDevices["Device_Index"]) == Convert.ToInt32(dictWanPorts["Device_Index"]))
                        {
                            //Console.WriteLine("Port on Device.");
                            //strPortName = Convert.ToString(dictWanPorts["Port_Name"]) + " (" + Convert.ToString(dictWanPorts["Port_Index"] + ")");
                            strPortName = Convert.ToString(dictWanPorts["Port_Name"]);
                            listShapesWanPorts.Add(page1.DrawRectangle(doubNextPortStartPointX, doubNextPortStartPointY, doubNextPortStartPointX + 0.6, doubNextPortStartPointY - 0.2));
                            listShapesWanPorts[listShapesWanPorts.Count - 1].Data1 = Convert.ToString(dictWanPorts["Port_Index"]);
                            listShapesWanPorts[listShapesWanPorts.Count - 1].Data2 = Convert.ToString(dictWanPorts["Link_Type"]);
                            listShapesWanPorts[listShapesWanPorts.Count - 1].Data3 = Convert.ToString(intCurrentWanChassis);
                            listShapesWanPorts[listShapesWanPorts.Count - 1].Text = strPortName;
                            doubNextPortStartPointX += 0.2;
                            listShapesWanPorts[listShapesWanPorts.Count - 1].Rotate90();
                            listShapesWanPorts[listShapesWanPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        };

                    };


                    doubWanLastX = doubDeviceEndPointX;
                    //intListCurrentIndex++;
                };

                

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                int intCurrentBypassPort;

                double doubUpperStartPoint = doubStartPointNextShapeY - 1.5;

                ////////////////////////////////////////////////////// Draw Bypass IS100 Chassis ////////////////////////////////////////////////////////////////

                List<Visio.Shape> listShapesBypass100Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass100NetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass100MonPorts = new List<Visio.Shape>();






                double doubBypassEndX;


                //doubStartPointNextShapeX = doubDeviceStartPointX + 8;

                doubStartPointNextShapeX = Math.Max(listShapesLanPorts.Count, listShapesWanPorts.Count);
                doubStartPointNextShapeY -= 1.5;


                intCurrentBypassPort = 0;

                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs100Bypasses; intCurrentBypassDevice++)
                {

                    listShapesBypass100Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 1, doubStartPointNextShapeY - 1.3));
                    listShapesBypass100Devices[listShapesBypass100Devices.Count - 1].Text = "IS100-" + intCurrentBypassDevice;
                    listShapesBypass100Devices[listShapesBypass100Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(255,228,225)";




                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    doubStartPointNextShapeY -= 2;





                    ///////////////////////////// Draw Bypass IS100 Ports ////////////////////////////////////

                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~    Original-Bypass ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 2; intCurrentPortCounterInBypass++)
                    {
                        intCurrentBypassPort++;
                        listShapesBypass100NetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3));
                        listShapesBypass100NetPorts[listShapesBypass100NetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass100NetPorts[listShapesBypass100NetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass100NetPorts[listShapesBypass100NetPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass100NetPorts[listShapesBypass100NetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listShapesBypass100MonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 1, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 1.4, doubNextPortStartPointY - 0.3));
                        listShapesBypass100MonPorts[listShapesBypass100MonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass100MonPorts[listShapesBypass100MonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass100MonPorts[listShapesBypass100MonPorts.Count - 1].Text = "Mon " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass100MonPorts[listShapesBypass100MonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrentBypassPort++;
                        listShapesBypass100NetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1));
                        listShapesBypass100NetPorts[listShapesBypass100NetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass100NetPorts[listShapesBypass100NetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass100NetPorts[listShapesBypass100NetPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/1";
                        listShapesBypass100NetPorts[listShapesBypass100NetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listShapesBypass100MonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 1, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 1.4, doubNextPortStartPointY - 0.1));
                        listShapesBypass100MonPorts[listShapesBypass100MonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass100MonPorts[listShapesBypass100MonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass100MonPorts[listShapesBypass100MonPorts.Count - 1].Text = "Mon " + intCurrentPortCounterInBypass + "/1";
                        listShapesBypass100MonPorts[listShapesBypass100MonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        doubNextPortStartPointY -= 0.7;
                    };




                }




                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                ////////////////////////////////////////////////////// Draw Bypass IS40 Chassis (for 40G) ////////////////////////////////////////////////////////////////

                List<Visio.Shape> listShapesBypass40Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass40NetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass40MonPorts = new List<Visio.Shape>();

                //doubStartPointNextShapeX = doubStartPointNextShapeX;
                //doubStartPointNextShapeY -= 1;





                intCurrentBypassPort = 0;

                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs40Bypasses; intCurrentBypassDevice++)
                {
                    listShapesBypass40Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 1, doubStartPointNextShapeY - 2));
                    listShapesBypass40Devices[listShapesBypass40Devices.Count - 1].Text = "IS40-" + intCurrentBypassDevice;
                    listShapesBypass40Devices[listShapesBypass40Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(255,165,0)";

                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    doubStartPointNextShapeY -= 2.7;





                    /////////////////////////// Draw Bypass IS40 Ports (40G) ////////////////////////////////////


                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 3; intCurrentPortCounterInBypass++)
                    {
                        intCurrentBypassPort++;
                        listShapesBypass40NetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3));
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listShapesBypass40MonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 1, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 1.4, doubNextPortStartPointY - 0.3));
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Text = "Mon " + intCurrentPortCounterInBypass + "/0";
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08"; 
                        intCurrentBypassPort++;
                        listShapesBypass40NetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1));
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/1";
                        listShapesBypass40NetPorts[listShapesBypass40NetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        listShapesBypass40MonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 1, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 1.4, doubNextPortStartPointY - 0.1));
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].Text = "Mon " + intCurrentPortCounterInBypass + "/1";
                        listShapesBypass40MonPorts[listShapesBypass40MonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        
                        doubNextPortStartPointY -= 0.7;
                    };

                }


                ////////////////////////////////////////////////////// Draw Bypass IS40 Chassis (for 10G) ////////////////////////////////////////////////////////////////

                List<Visio.Shape> listShapesBypass10Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass10NetPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass10MonPorts = new List<Visio.Shape>();

                List<Visio.Shape> listHydraLines = new List<Visio.Shape>();
                List<Visio.Shape> listBypassHydraConnectors = new List<Visio.Shape>();

                //doubStartPointNextShapeX = doubStartPointNextShapeX + 4;
                doubStartPointNextShapeY -= 1;

                intCurrentBypassPort = 0;

                //int intCountLinesInHydra;

                //double doubHydraFirstLineY;
                //double doubHydraLastLineY;

                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs10Bypasses; intCurrentBypassDevice++)
                {
                    listShapesBypass10Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 1, doubStartPointNextShapeY - 4.1));
                    listShapesBypass10Devices[listShapesBypass10Devices.Count - 1].Text = "IS40-" + intCurrentBypassDevice;
                    listShapesBypass10Devices[listShapesBypass10Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(219,112,147)";

                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    doubStartPointNextShapeY -= 5.4;





                    /////////////////////////// Draw Bypass IS40 Ports (for 10G) ////////////////////////////////////




                    //myList.Clear();
                    //intCountLinesInHydra = 0;


                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 3; intCurrentPortCounterInBypass++)
                    {
                        for (int intCurrentSubslotCounterInBypass = 1; intCurrentSubslotCounterInBypass <= 2; intCurrentSubslotCounterInBypass++)
                        {
                            intCurrentBypassPort++;
                            listShapesBypass10NetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3));
                            listShapesBypass10NetPorts[listShapesBypass10NetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                            listShapesBypass10NetPorts[listShapesBypass10NetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                            listShapesBypass10NetPorts[listShapesBypass10NetPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/0";
                            listShapesBypass10NetPorts[listShapesBypass10NetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            listShapesBypass10MonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 1, doubNextPortStartPointY - 0.5, doubNextPortStartPointX + 1.5, doubNextPortStartPointY - 0.3));
                            listShapesBypass10MonPorts[listShapesBypass10MonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                            listShapesBypass10MonPorts[listShapesBypass10MonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                            listShapesBypass10MonPorts[listShapesBypass10MonPorts.Count - 1].Text = "Mon " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/0";
                            listShapesBypass10MonPorts[listShapesBypass10MonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08"; 
                            intCurrentBypassPort++;
                            listShapesBypass10NetPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1));
                            listShapesBypass10NetPorts[listShapesBypass10NetPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                            listShapesBypass10NetPorts[listShapesBypass10NetPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                            listShapesBypass10NetPorts[listShapesBypass10NetPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/1";
                            listShapesBypass10NetPorts[listShapesBypass10NetPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                            listShapesBypass10MonPorts.Add(page1.DrawRectangle(doubNextPortStartPointX + 1, doubNextPortStartPointY - 0.3, doubNextPortStartPointX + 1.5, doubNextPortStartPointY - 0.1));
                            listShapesBypass10MonPorts[listShapesBypass10MonPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                            listShapesBypass10MonPorts[listShapesBypass10MonPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                            listShapesBypass10MonPorts[listShapesBypass10MonPorts.Count - 1].Text = "Mon " + intCurrentPortCounterInBypass + "/" + intCurrentSubslotCounterInBypass + "/1";
                            listShapesBypass10MonPorts[listShapesBypass10MonPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                            ///////////////     Draw Hydra Connector    //////////////////////////
                            if (!boolNoBalancer)
                            {
                                listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 1.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 1.5 + 0.1, doubNextPortStartPointY - 0.4));
                                listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 1.5, doubNextPortStartPointY - 0.2, doubNextPortStartPointX + 1.5 + 0.1, doubNextPortStartPointY - 0.2));

                                if (listHydraLines.Count == 4)
                                {
                                    listHydraLines.Add(page1.DrawLine(doubNextPortStartPointX + 1.5 + 0.1, doubNextPortStartPointY - 0.4, doubNextPortStartPointX + 1.5 + 0.1, doubNextPortStartPointY + 0.5));
                                    vsoWindow.DeselectAll();
                                    foreach (Visio.Shape objHydraSingleLine in listHydraLines)
                                    {
                                        vsoWindow.Select(objHydraSingleLine, 2);
                                    };

                                    vsoSelection = vsoWindow.Selection;
                                    listBypassHydraConnectors.Add(vsoSelection.Group());
                                    listBypassHydraConnectors[listBypassHydraConnectors.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);

                                    listHydraLines.Clear();
                                };

                            };


                            doubBypassEndX = doubStartPointNextShapeX + 1;



                            doubNextPortStartPointY -= 0.7;                 // Move down to next port
                        };

                    };

                };

                Console.WriteLine($"Total Bypass Hydra Count: {listBypassHydraConnectors.Count}");

                //foreach (Visio.Shape objLanPort in listShapesLanPorts)

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                ////////////////////////////////////////////////////// Draw Bypass IBS1U Chassis (1G Fiber Ports) ////////////////////////////////////////////////////////////////

                List<Visio.Shape> listShapesBypass1FiberDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass1FiberPorts = new List<Visio.Shape>();


                //doubStartPointNextShapeX = doubDeviceStartPointX + 4;
                doubStartPointNextShapeY -= 1.5;

                intCurrentBypassPort = 0;

                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs1FiberBypasses; intCurrentBypassDevice++)
                {
                    listShapesBypass1FiberDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 1, doubStartPointNextShapeY - 2.7));
                    listShapesBypass1FiberDevices[listShapesBypass1FiberDevices.Count - 1].Text = "IBS1U-" + intCurrentBypassDevice;
                    listShapesBypass1FiberDevices[listShapesBypass1FiberDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(255,228,225)";

                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    doubStartPointNextShapeY -= 3;





                    /////////////////////////// Draw Bypass IBS1U Ports (Fiber) ////////////////////////////////////


                    //doubBypassEndX

                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 4; intCurrentPortCounterInBypass++)
                    {
                        intCurrentBypassPort++;
                        listShapesBypass1FiberPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3));
                        listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0";
                        //listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0 (" + listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Data2 + ")";
                        listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrentBypassPort++;
                        listShapesBypass1FiberPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1));
                        listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/1";
                        //listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0 (" + listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Data2 + ")";
                        listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        doubNextPortStartPointY -= 0.7;
                    };






                }


                doubBypassEndX = doubStartPointNextShapeX + 1;
















                ////////////////////////////////////////////////////// Draw Bypass IBS1U Chassis (1G Copper Ports) ////////////////////////////////////////////////////////////////

                List<Visio.Shape> listShapesBypass1CopperDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBypass1CopperPorts = new List<Visio.Shape>();


                //doubStartPointNextShapeX = doubDeviceStartPointX + 4;
                doubStartPointNextShapeY -= 1.5;

                intCurrentBypassPort = 0;

                for (int intCurrentBypassDevice = 1; intCurrentBypassDevice <= intTotalIs1CopperBypasses; intCurrentBypassDevice++)
                {
                    listShapesBypass1CopperDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 1, doubStartPointNextShapeY - 2.7));
                    listShapesBypass1CopperDevices[listShapesBypass1CopperDevices.Count - 1].Text = "IBS1U-" + intCurrentBypassDevice;
                    listShapesBypass1CopperDevices[listShapesBypass1CopperDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(135,206,250)";

                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    doubStartPointNextShapeY -= 3;





                    /////////////////////////// Draw Bypass IBS1U Ports (Copper) ////////////////////////////////////

                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~    Original-Bypass ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



                    for (int intCurrentPortCounterInBypass = 1; intCurrentPortCounterInBypass <= 4; intCurrentPortCounterInBypass++)
                    {
                        intCurrentBypassPort++;
                        listShapesBypass1CopperPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.5, doubNextPortStartPointX, doubNextPortStartPointY - 0.3));
                        listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0";
                        //listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0 (" + listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Data2 + ")";
                        listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";
                        intCurrentBypassPort++;
                        listShapesBypass1CopperPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.4, doubNextPortStartPointY - 0.3, doubNextPortStartPointX, doubNextPortStartPointY - 0.1));
                        listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].Data1 = Convert.ToString(intCurrentBypassDevice);
                        listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].Data2 = Convert.ToString(intCurrentBypassPort);
                        listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/1";
                        //listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].Text = "Net " + intCurrentPortCounterInBypass + "/0 (" + listShapesBypass1FiberPorts[listShapesBypass1FiberPorts.Count - 1].Data2 + ")";
                        listShapesBypass1CopperPorts[listShapesBypass1CopperPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.08";

                        doubNextPortStartPointY -= 0.7;
                    };

                };

                doubBypassEndX = doubStartPointNextShapeX + 1;

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~






                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                ////////////////////    Draw LAN-Bypass Autoconnects    //////////////////////


                foreach (Visio.Shape objLanPort in listShapesLanPorts)
                {
                    if (objLanPort.Data2 == "100")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass100NetPorts)
                        {
                            if (Convert.ToInt32(objLanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objLanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };

                foreach (Visio.Shape objLanPort in listShapesLanPorts)
                {
                    if (objLanPort.Data2 == "40")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass40NetPorts)
                        {
                            if (Convert.ToInt32(objLanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objLanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };


                foreach (Visio.Shape objLanPort in listShapesLanPorts)
                {
                    if (objLanPort.Data2 == "10")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass10NetPorts)
                        {
                            if (Convert.ToInt32(objLanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objLanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };


                foreach (Visio.Shape objLanPort in listShapesLanPorts)
                {
                    if (objLanPort.Data2 == "1Fiber")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass1FiberPorts)
                        {
                            if (Convert.ToInt32(objLanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objLanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };



                foreach (Visio.Shape objLanPort in listShapesLanPorts)
                {
                    if (objLanPort.Data2 == "1Copper")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass1CopperPorts)
                        {
                            if (Convert.ToInt32(objLanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objLanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };













                ////////////////    Draw WAN-Bypass Autoconnects    //////////////////////

                foreach (Visio.Shape objWanPort in listShapesWanPorts)
                {
                    if (objWanPort.Data2 == "100")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass100NetPorts)
                        {
                            if (Convert.ToInt32(objWanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objWanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };


                foreach (Visio.Shape objWanPort in listShapesWanPorts)
                {
                    if (objWanPort.Data2 == "40")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass40NetPorts)
                        {
                            if (Convert.ToInt32(objWanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objWanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };


                foreach (Visio.Shape objWanPort in listShapesWanPorts)
                {
                    if (objWanPort.Data2 == "10")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass10NetPorts)
                        {
                            if (Convert.ToInt32(objWanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objWanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };

                foreach (Visio.Shape objWanPort in listShapesWanPorts)
                {
                    if (objWanPort.Data2 == "1Fiber")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass1FiberPorts)
                        {
                            if (Convert.ToInt32(objWanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objWanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };


                foreach (Visio.Shape objWanPort in listShapesWanPorts)
                {
                    if (objWanPort.Data2 == "1Copper")
                    {
                        foreach (Visio.Shape objBypassNetPort in listShapesBypass1CopperPorts)
                        {
                            if (Convert.ToInt32(objWanPort.Data1) == Convert.ToInt32(objBypassNetPort.Data2))
                                objWanPort.AutoConnect(objBypassNetPort, Visio.VisAutoConnectDir.visAutoConnectDirNone);
                        };
                    };

                };


                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw Balancers   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                double doubPortStartX;
                double doubPortEndX;
                double doubPortStartY;
                double doubPortEndY;

                List<Visio.Shape> listShapesBalancerDevices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBalancerUplinkPorts = new List<Visio.Shape>();
                List<Visio.Shape> listShapesBalancerDownlinkPorts = new List<Visio.Shape>();

                doubStartPointNextShapeX += 10;
                //doubStartPointNextShapeY = doubUpperStartPoint + 0.5;
                doubStartPointNextShapeY = doubUpperStartPoint;

                //doubStartPointNextShapeY += 8.7;

                //int intGlobalPortIndex = 0;

                //intBalancersQuantity

                //int intCurrentBalancerPort;

                //for (int intCurrentBalancerDevice = 1; intCurrentBalancerDevice <= intBalancersQuantity; intCurrentBalancerDevice++)
                //{

                int intCurrentBalancerDevice = 0;
                int intCurrentBalancerPort = 32;

                int intPortNumberAfterSwap;

                //doubNextPortStartPointX = doubStartPointNextShapeX;
                //doubNextPortStartPointY = doubStartPointNextShapeY;

                doubNextPortStartPointX = doubStartPointNextShapeX;
                doubNextPortStartPointY = doubStartPointNextShapeY;

                double doubShiftY;

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw 100G Ports ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                foreach (Visio.Shape objBypassMonPort in listShapesBypass100MonPorts)
                {
                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Drawing a New Balancer   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    //intCurrentBalancerPort++;
                    if (intCurrentBalancerPort == 32)
                    {
                        //Console.WriteLine($"Check");

                        intCurrentBalancerDevice++;
                        intCurrentBalancerPort = 16;
                        listShapesBalancerDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 9));
                        listShapesBalancerDevices[listShapesBalancerDevices.Count - 1].Text = "ELB-0133-" + intCurrentBalancerDevice;
                        listShapesBalancerDevices[listShapesBalancerDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(143,188,143)";
                       
                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw Balancer Downlink Ports ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                        doubNextPortStartPointY = doubStartPointNextShapeY + 7;

                        for (int intCurrentFloodedBalancerPort = 1; intCurrentFloodedBalancerPort <= 16; intCurrentFloodedBalancerPort++)
                        {
                            doubPortStartX = doubNextPortStartPointX;
                            doubPortStartY = doubNextPortStartPointY - 7 - 0.5;
                            doubPortEndX = doubNextPortStartPointX + 0.5;
                            doubPortEndY = doubNextPortStartPointY - 7 - 0.3;

                            listShapesBalancerDownlinkPorts.Add(page1.DrawRectangle(doubPortStartX, doubPortStartY, doubPortEndX, doubPortEndY));
                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].Data1 = Convert.ToString(intCurrentBalancerDevice);

                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].Text = "p" + intCurrentFloodedBalancerPort;
                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                            doubNextPortStartPointY -= 0.5;

                        };


                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;
                        doubStartPointNextShapeY -= 10;

                    };
                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Drawing 100G Ports on Balancer   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    intCurrentBalancerPort++;
                    //if (intCurrentBalancerPort == 33) continue;

                   // Console.Write($"Port Counter: {intCurrentBalancerPort}, Calculate: {intCurrentBalancerPort % 2}, ");
                    if (intCurrentBalancerPort % 2 == 0) intPortNumberAfterSwap = intCurrentBalancerPort - 1;
                    else intPortNumberAfterSwap = intCurrentBalancerPort + 1;
                    //Console.WriteLine($"Port No After Swap: {intPortNumberAfterSwap}");
                    //objBypassNetPort.Data2

                    if (intPortNumberAfterSwap % 2 == 0) doubShiftY = 0.2;
                    else doubShiftY = 0;


                    listShapesBalancerUplinkPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.5 - doubShiftY, doubNextPortStartPointX, doubNextPortStartPointY - 0.3 - doubShiftY));
                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Data1 = Convert.ToString(intCurrentBalancerDevice);
                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Data2 = objBypassMonPort.Data2;

                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Text = "p" + intPortNumberAfterSwap;
                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                    objBypassMonPort.AutoConnect(listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1], Visio.VisAutoConnectDir.visAutoConnectDirNone);

                    if (intPortNumberAfterSwap % 2 > 0) doubNextPortStartPointY -= 0.5;
                    //else doubNextPortStartPointY += 0.8;

                };

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
                        listShapesBalancerDevices[listShapesBalancerDevices.Count - 1].Text = "ELB-0133-" + intCurrentBalancerDevice;
                        listShapesBalancerDevices[listShapesBalancerDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(143,188,143)";

                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw Balancer Downlink Ports ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                        doubNextPortStartPointY = doubStartPointNextShapeY + 7;

                        for (int intCurrentFloodedBalancerPort = 1; intCurrentFloodedBalancerPort <= 16; intCurrentFloodedBalancerPort++)
                        {
                            doubPortStartX = doubNextPortStartPointX;
                            doubPortStartY = doubNextPortStartPointY - 7 - 0.5;
                            doubPortEndX = doubNextPortStartPointX + 0.5;
                            doubPortEndY = doubNextPortStartPointY - 7 - 0.3;

                            listShapesBalancerDownlinkPorts.Add(page1.DrawRectangle(doubPortStartX, doubPortStartY, doubPortEndX, doubPortEndY));
                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].Data1 = Convert.ToString(intCurrentBalancerDevice);

                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].Text = "p" + intCurrentFloodedBalancerPort;
                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

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




                //double doubHydraStartX;



                //Console.Write($"CurrentBalancerPorts: {intCurrentBalancerPort}, Calculate: {intCurrentBalancerPort % 2}, ");

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw 10G Ports ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                foreach (Visio.Shape objBypassMonHydra in listBypassHydraConnectors)
                {
                    //Console.WriteLine($"Check");
                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Drawing a New Balancer   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    //intCurrentBalancerPort++;
                    if (intCurrentBalancerPort == 32)
                    {
                        intCurrentBalancerDevice++;
                        intCurrentBalancerPort = 16;
                        listShapesBalancerDevices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 9));
                        listShapesBalancerDevices[listShapesBalancerDevices.Count - 1].Text = "ELB-0133-" + intCurrentBalancerDevice;
                        listShapesBalancerDevices[listShapesBalancerDevices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(143,188,143)";

                        //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw Balancer Downlink Ports ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        doubNextPortStartPointX = doubStartPointNextShapeX + 2;
                        doubNextPortStartPointY = doubStartPointNextShapeY + 7;

                        for (int intCurrentFloodedBalancerPort = 1; intCurrentFloodedBalancerPort <= 16; intCurrentFloodedBalancerPort++)
                        {
                            doubPortStartX = doubNextPortStartPointX;
                            doubPortStartY = doubNextPortStartPointY - 7 - 0.5;
                            doubPortEndX = doubNextPortStartPointX + 0.5;
                            doubPortEndY = doubNextPortStartPointY - 7 - 0.3;

                            listShapesBalancerDownlinkPorts.Add(page1.DrawRectangle(doubPortStartX, doubPortStartY, doubPortEndX, doubPortEndY));
                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].Data1 = Convert.ToString(intCurrentBalancerDevice);

                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].Text = "p" + intCurrentFloodedBalancerPort;
                            listShapesBalancerDownlinkPorts[listShapesBalancerDownlinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                            doubNextPortStartPointY -= 0.5;

                        };


                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;
                        doubStartPointNextShapeY -= 7;

                    };
                    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Drawing Ports on Balancer   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    intCurrentBalancerPort++;
                    //if (intCurrentBalancerPort == 33) continue;

                   // Console.Write($"Port Counter: {intCurrentBalancerPort}, Calculate: {intCurrentBalancerPort % 2}, ");
                    //if (intCurrentBalancerPort % 2 == 0) intPortNumberAfterSwap = intCurrentBalancerPort - 1;
                    //else intPortNumberAfterSwap = intCurrentBalancerPort + 1;
                    // Console.WriteLine($"Port No After Swap: {intPortNumberAfterSwap}");
                    //objBypassNetPort.Data2

                    //if (intPortNumberAfterSwap % 2 == 0) doubShiftY = 0.2;
                    //else doubShiftY = 0;
                    intPortNumberAfterSwap = intCurrentBalancerPort;
                    doubShiftY = 0;

                    doubPortStartX = doubNextPortStartPointX - 0.5;
                    doubPortStartY = doubNextPortStartPointY - 0.5 - doubShiftY;
                    doubPortEndX = doubNextPortStartPointX;
                    doubPortEndY = doubNextPortStartPointY - 0.3 - doubShiftY;



                    listShapesBalancerUplinkPorts.Add(page1.DrawRectangle(doubPortStartX, doubPortStartY, doubPortEndX, doubPortEndY));
                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Data1 = Convert.ToString(intCurrentBalancerDevice);
                    //listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Data2 = objBypassMonPort.Data2;

                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Text = "p" + intPortNumberAfterSwap;
                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";


                    objBypassMonHydra.AutoConnect(listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1], Visio.VisAutoConnectDir.visAutoConnectDirNone);


                    doubNextPortStartPointY -= 0.5;


                    //if (intPortNumberAfterSwap % 2 > 0) doubNextPortStartPointY -= 0.5;
                    //else doubNextPortStartPointY += 0.8;

                };

                //int intLastUsedBalancerPort = intCurrentBalancerPort;

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   Draw Rest of Ports on Last Balancer ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


                /// Draw Rest of Uplink Ports
                for (int intCurrentFloodedBalancerPort = intCurrentBalancerPort + 1; intCurrentFloodedBalancerPort <= 32; intCurrentFloodedBalancerPort++)
                {
                    doubPortStartX = doubNextPortStartPointX - 0.5;
                    doubPortStartY = doubNextPortStartPointY - 0.5;
                    doubPortEndX = doubNextPortStartPointX;
                    doubPortEndY = doubNextPortStartPointY - 0.3;

                    listShapesBalancerUplinkPorts.Add(page1.DrawRectangle(doubPortStartX, doubPortStartY, doubPortEndX, doubPortEndY));
                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Data1 = Convert.ToString(intCurrentBalancerDevice);

                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].Text = "p" + intCurrentFloodedBalancerPort;
                    listShapesBalancerUplinkPorts[listShapesBalancerUplinkPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                    doubNextPortStartPointY -= 0.5;
                }



                Console.WriteLine($"Balancers: {intCurrentBalancerDevice}");
                Console.WriteLine($"Balancers' Downlink Ports: {listShapesBalancerDownlinkPorts.Count}");

                /// Draw All Downplink Ports
                /// 

                //listShapesBalancerDownlinkPorts


                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////////////////////////////////////  Draw Filters ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                //intTotalFilters


                int intDownlinkPortOnBalancerToConnectFilterHydra;

                    doubStartPointNextShapeX += 12;
                
                doubStartPointNextShapeY = doubUpperStartPoint;



                List<Visio.Shape> listShapesFilter4160Devices = new List<Visio.Shape>();
                List<Visio.Shape> listShapesFilterPorts = new List<Visio.Shape>();
                List<Visio.Shape> listFilterHydraConnectors = new List<Visio.Shape>();

                List<Visio.Shape> listShapesSingleFilterPorts = new List<Visio.Shape>();


                // if (boolNoBalancer) Console.WriteLine("No Balancer");

                if (!boolNoBalancer)
                {
                    for (int intCurrentFilterDevice = 1; intCurrentFilterDevice <= intTotalFilters; intCurrentFilterDevice++)
                    {
                        listShapesFilter4160Devices.Add(page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 6));
                        listShapesFilter4160Devices[listShapesFilter4160Devices.Count - 1].Text = "Filter 4160-" + intCurrentFilterDevice;
                        listShapesFilter4160Devices[listShapesFilter4160Devices.Count - 1].get_Cells("FillForegnd").FormulaU = "=RGB(175,238,238)";

                        doubNextPortStartPointX = doubStartPointNextShapeX;
                        doubNextPortStartPointY = doubStartPointNextShapeY;

                        for (int intCurrentFilterPort = 1; intCurrentFilterPort <= 16; intCurrentFilterPort++)
                        {
                            listShapesFilterPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX, doubNextPortStartPointY - 0.2));
                            listShapesFilterPorts[listShapesFilterPorts.Count - 1].Text = "Te " + intCurrentFilterPort;
                            listShapesFilterPorts[listShapesFilterPorts.Count - 1].Data1 = Convert.ToString(intCurrentFilterDevice);
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
                                listFilterHydraConnectors[listFilterHydraConnectors.Count - 1].Data1 = Convert.ToString(intCurrentFilterDevice);
                                listHydraLines.Clear();

                                intDownlinkPortOnBalancerToConnectFilterHydra = (listFilterHydraConnectors.Count - 1 - (intCurrentFilterDevice - 1) * 4) * 16 + intCurrentFilterDevice - 1;

                                if (listShapesBalancerDownlinkPorts[intDownlinkPortOnBalancerToConnectFilterHydra] == null)
                                    intDownlinkPortOnBalancerToConnectFilterHydra = intDownlinkPortOnBalancerToConnectFilterHydra - 16 + intTotalFilters;


                                listFilterHydraConnectors[listFilterHydraConnectors.Count - 1].AutoConnect(listShapesBalancerDownlinkPorts[intDownlinkPortOnBalancerToConnectFilterHydra], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                                // 
                            };



                            if (intCurrentFilterPort % 2 > 0) doubNextPortStartPointY -= 0.2;
                            else doubNextPortStartPointY -= 0.5;

                        };

                        doubStartPointNextShapeY -= 7;
                    };
                }
                ////////////////////////////////    Working with Single-Filters without Balancers   //////////////////////////
                else
                {
                    doubStartPointNextShapeX = doubBypassEndX + 4;
                    doubStartPointNextShapeY -= 1;

                    Visio.Shape objStandAloneFilter = page1.DrawRectangle(doubStartPointNextShapeX, doubStartPointNextShapeY, doubStartPointNextShapeX + 2, doubStartPointNextShapeY - 0.4*intPortsNumberOnSingleFilter);
                    objStandAloneFilter.Text = strSingleFilterType;
                    objStandAloneFilter.get_Cells("FillForegnd").FormulaU = "=RGB(175,238,238)";

                    doubNextPortStartPointX = doubStartPointNextShapeX;
                    doubNextPortStartPointY = doubStartPointNextShapeY;

                    for (int intCurrentFilterPort = 1; intCurrentFilterPort <= intPortsNumberOnSingleFilter; intCurrentFilterPort++)
                    {
                        listShapesFilterPorts.Add(page1.DrawRectangle(doubNextPortStartPointX - 0.5, doubNextPortStartPointY - 0.4, doubNextPortStartPointX, doubNextPortStartPointY - 0.2));
                        listShapesFilterPorts[listShapesFilterPorts.Count - 1].Text = "Te " + intCurrentFilterPort;
                        listShapesFilterPorts[listShapesFilterPorts.Count - 1].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowFirst, (short)Visio.VisCellIndices.visCharacterSize).FormulaForceU = "0.12";

                        if (intCurrentFilterPort % 2 > 0) doubNextPortStartPointY -= 0.2;
                        else doubNextPortStartPointY -= 0.5;


                        //listShapesSingleFilterPorts
                    };

                    //Console.WriteLine($"Filter Ports: {listShapesFilterPorts.Count}");


                    int intSingleFilterPort = 0;
                    int intSwapPort;

                    foreach (Visio.Shape objBypassUplinkPort in listShapesBypass10MonPorts)
                    {
                        intSingleFilterPort++;
                        if (intSingleFilterPort % 2 == 0) intSwapPort = intSingleFilterPort - 1;
                        else intSwapPort = intSingleFilterPort + 1;

                        //Console.WriteLine($"Index: {intSwapPort - 1}");

                        if (intSwapPort <= listShapesFilterPorts.Count)
                        objBypassUplinkPort.AutoConnect(listShapesFilterPorts[intSwapPort - 1], Visio.VisAutoConnectDir.visAutoConnectDirNone);
                    };







                    //////  Group Standalone Filter & Ports ////////

                    vsoWindow.DeselectAll();
                    vsoWindow.Select(objStandAloneFilter, 2);
                    foreach (Visio.Shape objFilterPort in listShapesFilterPorts)
                    {
                        vsoWindow.Select(objFilterPort, 2);
                    };
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };


 



                //////////////////////////  Group Ports & Device    //////////////////////////////////////////////////////////////////////////////////////

                //////////////////////////  Group LAN Chassis & Ports    ////////////////////////////

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
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };

                //////////////////////////  Group 10G Bypass & Ports    ////////////////////////////
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
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };




                //////////////////////////  Group 100G Balancers & Ports    ////////////////////////////


                intCurrentDeviceInGroup = 0;


                foreach (Visio.Shape objSingleBalancer in listShapesBalancerDevices)
                {
                    vsoWindow.DeselectAll();
                    intCurrentDeviceInGroup++;
                    vsoWindow.Select(objSingleBalancer, 2);
                    foreach (Visio.Shape objBalancerPort in listShapesBalancerDownlinkPorts)
                    {
                        if (Convert.ToInt32(objBalancerPort.Data1) == intCurrentDeviceInGroup)
                        {
                            vsoWindow.Select(objBalancerPort, 2);
                        };

                    };
                    foreach (Visio.Shape objBalancerPort in listShapesBalancerUplinkPorts)
                    {
                        if (Convert.ToInt32(objBalancerPort.Data1) == intCurrentDeviceInGroup)
                        {
                            vsoWindow.Select(objBalancerPort, 2);
                        };

                    };
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                };



                //////////////////////////  Group 100G Filters & Ports    ////////////////////////////

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
                    
                    
                    vsoSelection = vsoWindow.Selection;
                    vsoSelection.Group();
                    
                };
            }

                //////////////////////////  Group Single Filter & Ports    ////////////////////////////





                //////////////////////////  Template    /////////////////////////////////////

                /*

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





                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////



                //Console Output
                //Console.WriteLine("Check");

                //Console.WriteLine(strSelectedPath);
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