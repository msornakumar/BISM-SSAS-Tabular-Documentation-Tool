using Microsoft.Office.Interop.Excel;
using System;
using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
using System.Windows.Forms;



namespace BISMDocumenterTOM
{
    public class BISMDocumenterCls
    {
        string AppPath;
        string OutputPath;

        Microsoft.Office.Interop.Excel.Application XLApp;
        Microsoft.Office.Interop.Excel.Workbook XLWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet XLWorkSheet;

        String OLAPServerName;
        String OLAPDBName;
        String OLAPCubeName;
        
        AMO.Server OLAPServer;
        AMO.Database OLAPDatabase;
        AMO.Cube OLAPCube;
        

        TOM.Server TOMServer;
        TOM.Database TOMDb;

        string Progressstr;

        //string  txtProgress;

        
        int ExcelSheetStartrow;

        int ExcelColNumber = 0;

        BISMDocumenterLibrary.ProgressWritter PX = new BISMDocumenterLibrary.ProgressWritter();
        
       

        public void GenerateDocument(String ServerName, String DBName, String CubeName, String DocumentPath, String FileName,System.Windows.Forms.TextBox progressTextBox,Boolean OpenXl)
        {
            try
            {
                String ConnStr;
                OLAPServerName = ServerName;
                // txtProgress= "";

                ConnStr = "Provider=MSOLAP;Data Source=" + OLAPServerName + ";";
                //Initial Catalog=Adventure Works DW 2008R2;"; 
                //OLAPServer = new AMO.Server();
                //OLAPServer.Connect(ConnStr);

                TOMServer = new TOM.Server();
                TOMServer.Connect(ConnStr);

            }
            catch (Exception err)
            {
                string errormsg = err.InnerException.ToString();
                // txtProgress = // txtProgress + "--------------------------------------------------------------------------------------" + Environment.NewLine;
                // txtProgress = // txtProgress + "Error Occured" + Environment.NewLine;
                // txtProgress = // txtProgress + err.InnerException.ToString() + Environment.NewLine;
            }

            try
            {

                Progressstr = "Generation started....";
                PX.InvokedAppType = "Windows";
                PX.WriteProgress(Progressstr, progressTextBox);

               

                //ProgressWritter.WriteProgress()
                


                if (!System.IO.Directory.Exists(DocumentPath))
                {
                    System.IO.Directory.CreateDirectory(DocumentPath);
                }

                OutputPath = DocumentPath;


                OLAPDBName = DBName;
                OLAPCubeName = CubeName;

                //OLAPDatabase = OLAPServer.Databases[OLAPDBName.Trim()];

                int TOMCompatibilityLevel;
                
                if (OLAPDBName.Trim() != "")
                {
                    TOMDb = TOMServer.Databases[OLAPDBName.Trim()];
                }
                else
                {
                    TOMDb = TOMServer.Databases[0];
                    OLAPDBName = TOMDb.Name;
                    OLAPCubeName = CubeName;

                }

                TOMCompatibilityLevel = TOMDb.CompatibilityLevel;

                Progressstr = "Database Compatibility Level - " + TOMDb.CompatibilityLevel.ToString();
                PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = "TOM Extracting Metadata for " + ServerName + " - " + DBName;
                PX.WriteProgress(Progressstr, progressTextBox);



                    FileName = FileName + ".xlsx";
                
                XLApp = new Microsoft.Office.Interop.Excel.Application();
                XLApp.Visible = false;

                XLApp.DisplayAlerts = false;
                XLWorkBook = XLApp.Workbooks.Add();
                XLWorkBook.SaveAs(OutputPath + "\\" + FileName);
                
                XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add();
                XLWorkSheet.Name = "Server";

                WriteHeaderCell(XLWorkSheet, 1, 1, "Server Name");
                XLWorkSheet.Cells[1, 2] = ServerName;
                FormatCell(XLWorkSheet, 1, 2, -1);

                WriteHeaderCell(XLWorkSheet, 2, 1, "Database Name");
                XLWorkSheet.Cells[2, 2] = OLAPDBName;
                FormatCell(XLWorkSheet, 2, 2, -1);

                WriteHeaderCell(XLWorkSheet, 3, 1, "Database CompatibilityLevel");
                XLWorkSheet.Cells[3, 2] = TOMCompatibilityLevel;
                FormatCell(XLWorkSheet, 3, 2, -1);

                //Progressstr = "TOM Extracting Metadata for " + OLAPDatabase + " - " + OLAPCube;

                XLWorkBook.Save();

                String ProgressStringStartTemplate = "Generating Documentation for <PlaceHolder>....";

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>","Connections");
                PX.WriteProgress(Progressstr, progressTextBox);
                
                TOMExtractConnections();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);


                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Dimension");
                PX.WriteProgress(Progressstr, progressTextBox);

                TOMExtractDimension();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Dimension Attributes");
                PX.WriteProgress(Progressstr, progressTextBox);

                TOMExtractDimensionAttribute();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);


                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Relationships");
                PX.WriteProgress(Progressstr, progressTextBox);

                TOMExtractRelationship();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Hierarchies");
                PX.WriteProgress(Progressstr, progressTextBox);

                TOMExtractHierarchies();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Measures");
                PX.WriteProgress(Progressstr, progressTextBox);

                TOMExtractMeasures();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "KPIs");
                PX.WriteProgress(Progressstr, progressTextBox);

                TOMExtractKPIs();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Partitions");
                PX.WriteProgress(Progressstr, progressTextBox);

                TOMExtractPartitions();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Perspectives");
                PX.WriteProgress(Progressstr, progressTextBox);

                TOMExtractPerspectives();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Roles");
                PX.WriteProgress(Progressstr, progressTextBox);

                TOMExtractRole();
                FormatSheet(XLWorkSheet);

                Progressstr = Progressstr.Replace("....", " Completed");
                PX.WriteProgress(Progressstr, progressTextBox);



                bool sheet1exists = false;
                bool sheet2exists = false;
                bool sheet3exists = false;


                foreach (Worksheet sheet in XLWorkBook.Sheets)
                {
                    // Check the name of the current sheet
                    if (sheet.Name == "Sheet1")
                    {

                        sheet1exists = true;


                    }

                    if (sheet.Name == "Sheet2")
                    {
                        sheet2exists = true;

                    }

                    if (sheet.Name == "Sheet3")
                    {
                        sheet3exists = true;

                    }
                }

                
                
                if (sheet1exists == true)
                {
                    XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet) XLWorkBook.Sheets["Sheet1"];
                    XLWorkSheet.Delete();

                }

                if (sheet2exists == true)
                {
                    XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Sheet2"];
                    XLWorkSheet.Delete();

                }

                if (sheet3exists == true)
                {
                    XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Sheet3"];
                    XLWorkSheet.Delete();
                }

                XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Server"];
                XLWorkSheet.Activate();



                // txtProgress.AppendText(Progressstr + " Completed " + Environment.NewLine);
                //      }
                //  }

                if (OpenXl == false)
                {
                    XLWorkBook.Close(true);
                    XLApp.Quit();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XLApp);
                }
                else
                {
                    XLApp.Visible = true;
                    XLApp.WindowState = XlWindowState.xlMaximized;
                }


            }
            catch (Exception err)
            {

                string errormsg = err.ToString();
                Progressstr = "--------------------------------------------------------------------------------------" + Environment.NewLine ;
                Progressstr = Progressstr + "Error Occured" + Environment.NewLine;
                Progressstr = Progressstr + "--------------------------------------------------------------------------------------" + Environment.NewLine;
                Progressstr = Progressstr + errormsg;
                // MessageBox.Show(errormsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                PX.WriteProgress(Progressstr, progressTextBox);

                if (XLWorkBook != null)
                {

                    XLWorkBook.Save();
                    XLWorkBook.Close(true);


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkBook);

                }

                if (XLApp != null)
                {
                    XLApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XLApp);
                }

            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLApp);
                GC.Collect();
            }


        }

        public void TOMExtractConnections()
        {
            // txtProgress.AppendText(Progressstr + " Connections Started " + Environment.NewLine);


            ExcelSheetStartrow = 6;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, 1, "Connections");

            ExcelSheetStartrow++;
            ExcelColNumber = 1;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Connection Name");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Connection String");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Description");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Protocol");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Server");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Schema");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Database");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Domain");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Account");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "EmailAddress");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Path");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Url");
            /*
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Connection Name");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Connection String");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Description");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Server");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Database");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Protocol");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Account");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Domain");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "EmailAddress");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Path");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Property");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Resource");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Schema");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "Url");
            ColNumber++;
            WriteHeaderCell(XLWorkSheet, 6, ColNumber, "View");
            */

            ExcelSheetStartrow++; 
            
            TOM.ProviderDataSource TOMProviderDs;
            TOM.StructuredDataSource TOMStructuredDs;

            ExcelColNumber = 1;
            // for (int I = 0; I <= TOMDb.Model.DataSources.Count - 1; I++)
            foreach(TOM.DataSource TomDatasource in TOMDb.Model.DataSources)
            {

               
               
                if (TomDatasource.Type == TOM.DataSourceType.Provider)
                {
                    //Connection Name
                    TOMProviderDs = (TOM.ProviderDataSource)TomDatasource;
                    ExcelColNumber = 1;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMProviderDs.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    
                    //Connection String
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMProviderDs.ConnectionString;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMProviderDs.Description;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                }
                else if(TomDatasource.Type == TOM.DataSourceType.Structured)
                {
                    TOMStructuredDs = (TOM.StructuredDataSource)TomDatasource;

                    //Connection Name
                    ExcelColNumber = 1;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Connection String
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Address.ConnectionString;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Description
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.Description;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Protocol
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Protocol;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    //Server
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Address.Server;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    //Schema
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Address.Schema;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    //Database
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Address.Database;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    //Domain
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Address.Domain;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    //Account
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Address.Account;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    //EmailAddress
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Address.EmailAddress;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    //Path
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Address.Path;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    //Url
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMStructuredDs.ConnectionDetails.Address.Url;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);



                }



                /*
                  ExcelColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMDs.Provider;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                 *
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.Database;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Protocol;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.Account;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.Domain;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.EmailAddress;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.Path;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.Property;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.Resource;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.Schema;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.Url;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.View;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                */

                ExcelSheetStartrow++;
            }

            /*
            foreach (TOM.ProviderDataSource pds in TOMDb.Model.DataSources)
            {
               
                    XLWorkSheet.Cells[ExcelSheetStartrow, 1] = pds.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 2] = pds.ConnectionString;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 3] = pds.Provider;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 4] = pds.Description;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);

                ExcelSheetStartrow++;
              
            }
            */
            XLWorkBook.Save();
            // txtProgress.AppendText(Progressstr + " Connections Completed " + Environment.NewLine);
        }

        public void TOMExtractDimension()
        {
            // try
            // {
            // txtProgress.AppendText(Progressstr + " Dimensions Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;


            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Server"]);
            XLWorkSheet.Name = "Dimensions";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Dimensions"];

            ExcelColNumber = 1;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "DimensionName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Mode");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "SourceType");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Data Source");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Source Query \\ Expression");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Source Schema Name");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Source Table Name");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "isHidden");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Category");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Description");






            foreach (TOM.Table Dimension in TOMDb.Model.Tables)
            {

                //DimensionName
                ExcelColNumber = 1;
                XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = Dimension.Name;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                //Mode
                ExcelColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = Dimension.Partitions[0].Mode.ToString();
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                //SourceType
                ExcelColNumber++;
                String DimensionSourceType = Dimension.Partitions[0].SourceType.ToString();
                XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimensionSourceType;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);


                

                TOM.QueryPartitionSource TOMPartitionSource;
                if (DimensionSourceType == TOM.PartitionSourceType.Query.ToString())
                {
                    TOMPartitionSource = (TOM.QueryPartitionSource)Dimension.Partitions[0].Source;

                    //Data Source
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartitionSource.DataSource.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    // Source Query \\ Expression
                    ExcelColNumber++;
                    if (Dimension.Annotations.Find("_TM_ExtProp_QueryDefinition") != null)
                    {

                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = Dimension.Annotations["_TM_ExtProp_QueryDefinition"].Value;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    }

                    
                }

                TOM.CalculatedPartitionSource TOMCalcPartitionSource;
                if (DimensionSourceType == TOM.PartitionSourceType.Calculated.ToString())
                {
                    TOMCalcPartitionSource = (TOM.CalculatedPartitionSource)Dimension.Partitions[0].Source;

                    //Data Source
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = "";
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    // Source Query \\ Expression
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMCalcPartitionSource.Expression ;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                }

                TOM.MPartitionSource TOMMPartitionSource;
                if (DimensionSourceType == TOM.PartitionSourceType.M.ToString())
                {
                    TOMMPartitionSource = (TOM.MPartitionSource)Dimension.Partitions[0].Source;

                    //Data Source
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = "";
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    // Source Query \\ Expression
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMPartitionSource.Expression;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                }

                TOM.EntityPartitionSource TOMEntitiySource;
                if (DimensionSourceType == TOM.PartitionSourceType.Entity.ToString())
                {
                    TOMEntitiySource = (TOM.EntityPartitionSource)Dimension.Partitions[0].Source;

                    //Data Source
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMEntitiySource.DataSource.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    // Source Query \\ Expression
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMEntitiySource.EntityName;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                }

                // To handle case of None
                if (DimensionSourceType == TOM.PartitionSourceType.None.ToString())
                {
                    ExcelColNumber++;
                    ExcelColNumber++;
                }

                // Source Schema Name
                ExcelColNumber++;
                if (Dimension.Annotations.Find("_TM_ExtProp_DbSchemaName") != null)
                {
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = Dimension.Annotations["_TM_ExtProp_DbSchemaName"].Value;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                }

                // Source Table Name
                ExcelColNumber++;
                if (Dimension.Annotations.Find("_TM_ExtProp_DbTableName") != null)
                {

                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = Dimension.Annotations["_TM_ExtProp_DbTableName"].Value;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                }
                // isHidden
                ExcelColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = Dimension.IsHidden;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                //Category
                ExcelColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = Dimension.DataCategory;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);


                // Description
                ExcelColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = Dimension.Description;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                ExcelSheetStartrow++;

            }
            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Dimensions Completed " + Environment.NewLine);
            /*
             }
            catch (Exception err)
            {

                string errormsg = err.ToString();
                // txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                // txtProgress.AppendText("Error Occured" + Environment.NewLine);
                // txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                // txtProgress.AppendText(errormsg + Environment.NewLine);

                throw (err);

            }
             * */

        }

        public void TOMExtractDimensionAttribute()
        {
            // txtProgress.AppendText(Progressstr + " Dimension Attributes Started " + Environment.NewLine);
            string ColumnSource;
            ExcelSheetStartrow = 2;
            // XLWorkSheet = XLWorkBook.Sheets["DimensionAttributes"];
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Dimensions"]);
            XLWorkSheet.Name = "DimensionAttributes";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["DimensionAttributes"];

            ExcelColNumber = 1;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Dimension Name");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Attribute Name");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Column Type");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Description");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Data Type");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Source DB ColumnName \\ Expression");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Source Data Type");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Alignment");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "DataCategory");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "DisplayFolder");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "DisplayOrdinal");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "EncodingHint");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "FormatString");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "IsAvailableInMDX");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "IsHidden");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "IsUnique");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "KeepUniqueRows");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "SortByColumn");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "SummarizeBy");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "TableDetailPosition");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Default Label");

            foreach (TOM.Table Dimension in TOMDb.Model.Tables)
            {
                foreach (TOM.Column DimAttribute in Dimension.Columns)
                {

                    
                    

                    //"Dimension Name"
                    ExcelColNumber = 1;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = Dimension.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Attribute Name
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Column Type
                    ExcelColNumber++;
                    String ColumnType = DimAttribute.Type.ToString();
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = ColumnType;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);


                    //Description
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.Description;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Data Type
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.DataType.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    TOM.DataColumn TomDataColumn ;
                    TOM.RowNumberColumn TOMRowNumberColumn;
                    TOM.CalculatedTableColumn TOMCalculatedTableColumn;
                    TOM.CalculatedColumn TOMCalculatedColumn;

                    if (ColumnType == TOM.ColumnType.Data.ToString())
                    {
                        TomDataColumn = (TOM.DataColumn)DimAttribute;

                        //Source DB ColumnName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TomDataColumn.SourceColumn;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //Source Data Type
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TomDataColumn.SourceProviderType;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    }

                    if (ColumnType == TOM.ColumnType.RowNumber.ToString())
                    {
                        TOMRowNumberColumn = (TOM.RowNumberColumn)DimAttribute;

                        //Source DB ColumnName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMRowNumberColumn.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //Source Data Type
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMRowNumberColumn.SourceProviderType;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    }

                    if (ColumnType == TOM.ColumnType.CalculatedTableColumn.ToString())
                    {
                        TOMCalculatedTableColumn = (TOM.CalculatedTableColumn)DimAttribute;

                        
                        //Source DB ColumnName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMCalculatedTableColumn.SourceColumn  ;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //Source Data Type
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMCalculatedTableColumn.SourceProviderType;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    }

                    if (ColumnType == TOM.ColumnType.Calculated.ToString())
                    {
                        TOMCalculatedColumn = (TOM.CalculatedColumn)DimAttribute;


                        //Source DB ColumnName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMCalculatedColumn.Expression;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //Source Data Type
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMCalculatedColumn.SourceProviderType;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    }

                    //Alignment
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.Alignment.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //DataCategory
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.DataCategory;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //DisplayFolder
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.DisplayFolder;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //DisplayOrdinal
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.DisplayOrdinal.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //EncodingHint
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.EncodingHint.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //FormatString
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.FormatString;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //IsAvailableInMDX
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.IsAvailableInMDX.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    
                    //IsHidden
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.IsHidden.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    
                    //IsUnique
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.IsUnique.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    
                    //KeepUniqueRows
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.KeepUniqueRows.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    
                    //SortByColumn
                    ExcelColNumber++;

                    if (DimAttribute.SortByColumn != null)
                    {
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.SortByColumn.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    }
                    //SummarizeBy
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.SummarizeBy.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    
                    
                    //TableDetailPosition
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.TableDetailPosition.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = DimAttribute.IsDefaultLabel.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);


                    ExcelSheetStartrow++;
                   
                }
            }
            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Dimension Attributes Completed " + Environment.NewLine);

        }

        public void TOMExtractRelationship()
        {
            // txtProgress.AppendText(Progressstr + " Relationships Started " + Environment.NewLine);
            ExcelSheetStartrow = 1;
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["DimensionAttributes"]);
            
            XLWorkSheet.Name = "Relationships";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Relationships"];

            ExcelColNumber=1;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "FromTable");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "FromColumn");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "FromCardinality");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "ToTable");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "ToColumn");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "ToCardinality");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "IsActive");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "CrossFilteringBehavior");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "JoinOnDateBehavior");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "RelyOnReferentialIntegrity");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "SecurityFilteringBehavior");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, 1, ExcelColNumber, "Type");



            ExcelSheetStartrow++;

            TOM.SingleColumnRelationship TOMsingleColumnRelationship;

            foreach(TOM.Relationship TOMRels in TOMDb.Model.Relationships )
            {

                if (TOMRels.Type == TOM.RelationshipType.SingleColumn)
                {
                    TOMsingleColumnRelationship = (TOM.SingleColumnRelationship)TOMRels;

                    ExcelColNumber=1;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.FromTable.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.FromColumn.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.FromCardinality.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.ToTable.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.ToColumn.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.ToCardinality.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.IsActive.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.CrossFilteringBehavior.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.JoinOnDateBehavior.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.RelyOnReferentialIntegrity.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.SecurityFilteringBehavior.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMsingleColumnRelationship.Type.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                }

                ExcelSheetStartrow++;

            }

            /*
            foreach (AMO.Dimension RelDimension in OLAPDatabase.Dimensions)
            {
                foreach (AMO.Relationship DimRelationship in RelDimension.Relationships)
                {
                    XLWorkSheet.Cells[ExcelSheetStartrow, 1] = RelDimension.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                    foreach (AMO.RelationshipEndAttribute FromRelAttribute in DimRelationship.FromRelationshipEnd.Attributes)
                    {
                        XLWorkSheet.Cells[ExcelSheetStartrow, 2] = RelDimension.Attributes[FromRelAttribute.AttributeID.ToString()].Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 3] = DimRelationship.FromRelationshipEnd.Multiplicity;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                    }
                    foreach (AMO.RelationshipEndAttribute ToRelAttribute in DimRelationship.ToRelationshipEnd.Attributes)
                    {

                        XLWorkSheet.Cells[ExcelSheetStartrow, 4] = OLAPCube.Dimensions.Find(DimRelationship.ToRelationshipEnd.DimensionID).Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 5] = OLAPCube.Dimensions.Find(DimRelationship.ToRelationshipEnd.DimensionID).Attributes.Find(ToRelAttribute.AttributeID);
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 5, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 6] = DimRelationship.ToRelationshipEnd.Multiplicity;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 6, -1);
                    }

                    ExcelSheetStartrow++;
                }



            }

            */

            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Relationships Completed " + Environment.NewLine);

        }

        public void TOMExtractHierarchies()
        {
            // txtProgress.AppendText(Progressstr + " Hierarchies Started " + Environment.NewLine);
            ExcelSheetStartrow = 1;

          

             XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Relationships"]);
            //XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add();
            XLWorkSheet.Name = "Hierarchies";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Hierarchies"];


            ExcelColNumber=1;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "HierarchyTableName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "HierarchyName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Description");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "DisplayFolder");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "HideMembers");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "IsHidden");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "LevelOrdinal");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "LevelName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "LevelColumnName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "LevelDescription");


            ExcelSheetStartrow++;

            
            foreach (TOM.Table TOMTable in TOMDb.Model.Tables)
            {
                foreach(TOM.Hierarchy TOMHierarchy in TOMTable.Hierarchies)
                {

                    foreach (TOM.Level TOMHierarchyLevel in TOMHierarchy.Levels)
                    {
                        ExcelColNumber = 1;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchy.Table.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchy.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchy.Description;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchy.DisplayFolder;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchy.HideMembers.ToString();
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchy.IsHidden.ToString();
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchyLevel.Ordinal;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchyLevel.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchyLevel.Column.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMHierarchyLevel.Description;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelSheetStartrow++;
                    }
                    
                }

            }
            









            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Hierarchies Completed " + Environment.NewLine);

        }

        public void TOMExtractMeasures()
        {
            // txtProgress.AppendText(Progressstr + " Measures Started " + Environment.NewLine);
            ExcelSheetStartrow = 1;


            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Hierarchies"]);
            XLWorkSheet.Name = "Measures";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Measures"];


            

            ExcelColNumber=1;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Name");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "DataType");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Expression");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "FormatString");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Measure Table");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "DisplayFolder");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Description");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "DetailRowsDefinition");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "IsHidden");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "IsSimpleMeasure");

            ExcelSheetStartrow++;

            foreach (TOM.Table TOMTable in TOMDb.Model.Tables)
            {
                foreach (TOM.Measure TOMMeasure in TOMTable.Measures)
                {
                    ExcelColNumber=1;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMeasure.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMeasure.DataType.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMeasure.Expression;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMeasure.FormatString;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMTable.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMeasure.DisplayFolder;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMeasure.Description;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMeasure.DetailRowsDefinition;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMeasure.IsHidden.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMeasure.IsSimpleMeasure.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    ExcelSheetStartrow++;
                }
            }




            
            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Measures Completed " + Environment.NewLine);

        }

        public void TOMExtractKPIs()
        {
            // txtProgress.AppendText(Progressstr + " KPIs Started " + Environment.NewLine);
            ExcelSheetStartrow = 1;

            

            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Measures"]);
            XLWorkSheet.Name = "KPIs";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["KPIs"];

            ExcelColNumber=1;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "KPIMeasure");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "KPIDescription");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "StatusExpression");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "StatusGraphic");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "StatusDescription");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "TargetExpression");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "TargetFormatString");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "TargetDescription");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "TrendExpression");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "TrendGraphic");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "TrendDescription");


            ExcelSheetStartrow++;


            TOM.KPI TOMKPI;
            foreach (TOM.Table TOMTable in TOMDb.Model.Tables)
            {
                foreach (TOM.Measure TOMMeasure in TOMTable.Measures)
                {
                    TOMKPI = TOMMeasure.KPI;

                    if (TOMKPI != null)
                    {

                        ExcelColNumber = 1;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.Measure.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.Description;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.StatusExpression;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.StatusGraphic;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.StatusDescription;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.TargetExpression;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.TargetFormatString;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.TargetDescription;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.TrendExpression;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.TrendGraphic;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMKPI.TrendDescription;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelSheetStartrow++;
                    }
                }
            }

            
            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " KPIs Completed " + Environment.NewLine);

        }



        public void TOMExtractPartitions()
        {
            // txtProgress.AppendText(Progressstr + " Partitions Started " + Environment.NewLine);
            ExcelSheetStartrow = 1;

            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["KPIs"]);
            XLWorkSheet.Name = "Partitions";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Partitions"];

           
            ExcelColNumber=1;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "PartitionName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "TableName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "SourceType");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Query \\ Expression");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Data Source");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Mode");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "DataView");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Description");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "RetainDataTillForceCalculate");

            ExcelSheetStartrow++;



            foreach (TOM.Table TOMTable in TOMDb.Model.Tables)
            {
                foreach(TOM.Partition TOMPartition in TOMTable.Partitions)
                {
                    //PartitionName
                    ExcelColNumber=1;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartition.Name; ;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //TableName
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartition.Table.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //SourceType
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartition.SourceType.ToString(); ;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    TOM.QueryPartitionSource TOMPartitionSource;
                    if (TOMPartition.SourceType.ToString() == TOM.PartitionSourceType.Query.ToString())
                    {
                        TOMPartitionSource = (TOM.QueryPartitionSource)TOMPartition.Source;

                        //Query \\ Expression
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartitionSource.Query.ToString();
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        
                        //Data Source
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartitionSource.DataSource.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        

                    }

                    TOM.CalculatedPartitionSource TOMCalcPartitionSource;
                    if (TOMPartition.SourceType.ToString() == TOM.PartitionSourceType.Calculated.ToString())
                    {
                        TOMCalcPartitionSource = (TOM.CalculatedPartitionSource)TOMPartition.Source;

                        //Query \\ Expression
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMCalcPartitionSource.Expression;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        
                        //Data Source
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = "";
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        
                    }

                    TOM.MPartitionSource TOMMPartitionSource;
                    if (TOMPartition.SourceType.ToString() == TOM.PartitionSourceType.M.ToString())
                    {
                        TOMMPartitionSource = (TOM.MPartitionSource)TOMPartition.Source;

                        //Query \\ Expression
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMMPartitionSource.Expression;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        
                        //Data Source
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = "";
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        
                    }

                    TOM.EntityPartitionSource TOMEntitiySource;
                    if (TOMPartition.SourceType.ToString() == TOM.PartitionSourceType.Entity.ToString())
                    {
                        TOMEntitiySource = (TOM.EntityPartitionSource)TOMPartition.Source;

                        //Query \\ Expression
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMEntitiySource.EntityName;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);
                        
                        //Data Source
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMEntitiySource.DataSource.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        
                    }

                    //Mode
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartition.Mode.ToString(); ;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Data View
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartition.DataView.ToString(); ;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    // Description
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartition.Description; ;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //RetainDataTillForceCalculate
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMPartition.RetainDataTillForceCalculate.ToString(); ;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    ExcelSheetStartrow++;
                }
            }

            

            

            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Partitions Completed " + Environment.NewLine);

        }

        public void TOMExtractPerspectives()
        {
            // txtProgress.AppendText(Progressstr + " Perspectives Started " + Environment.NewLine);
            ExcelSheetStartrow = 1;

            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Partitions"]);
            XLWorkSheet.Name = "Perspectives";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Perspectives"];


            ExcelColNumber = 1;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "PerspectiveName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "TableName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "ObjectName");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "ObjectType");

            ExcelSheetStartrow++;

            String PerspectiveName = "";
            String PerspectiveTableName = "";

            foreach (TOM.Perspective TOMPerspective in TOMDb.Model.Perspectives)
            {
                PerspectiveName = TOMPerspective.Name;
                foreach (TOM.PerspectiveTable TOMPerspectiveTable in TOMPerspective.PerspectiveTables)
                {
                    PerspectiveTableName = TOMPerspectiveTable.Name;
                    foreach (TOM.PerspectiveColumn PerspectiveColumn in TOMPerspectiveTable.PerspectiveColumns)
                    {

                        //PerspectiveName
                        ExcelColNumber = 1;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveName ;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //TableName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveTableName;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //ObjectName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveColumn.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //ObjectType
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveColumn.ObjectType.ToString();
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelSheetStartrow++;
                    }

                    foreach (TOM.PerspectiveMeasure PerspectiveMeasure in TOMPerspectiveTable.PerspectiveMeasures)
                    {

                        //PerspectiveName
                        ExcelColNumber = 1;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveName;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //TableName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveTableName;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //ObjectName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveMeasure.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //ObjectType
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveMeasure.ObjectType.ToString();
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelSheetStartrow++;
                    }

                    foreach (TOM.PerspectiveHierarchy PerspectiveHierarchy in TOMPerspectiveTable.PerspectiveHierarchies)
                    {

                        //PerspectiveName
                        ExcelColNumber = 1;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveName;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //TableName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveTableName;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //ObjectName
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveHierarchy.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        //ObjectType
                        ExcelColNumber++;
                        XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = PerspectiveHierarchy.ObjectType.ToString();
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                        ExcelSheetStartrow++;
                    }
                }
            }

            
            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Perspectives Completed " + Environment.NewLine);
        }

        public void TOMExtractRole()
        {
            // txtProgress.AppendText(Progressstr + " Roles Started " + Environment.NewLine);
            ExcelSheetStartrow = 1;

            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Perspectives"]);

            XLWorkSheet.Name = "Roles";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Roles"];

            ExcelColNumber = 1;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Role Name");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Role Permission");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Role Description");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Role Members");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Table");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "Filter Expression");
            ExcelColNumber++;
            WriteHeaderCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, "MetadataPermission");

            ExcelSheetStartrow++;

            String TOMModelRoleName;
            String TOMModelRolePermission;
            String TOMModelRoleDescription;
            String TOMModelRoleMembers;

            foreach (TOM.ModelRole TOMModelRole in TOMDb.Model.Roles )
            {
                TOMModelRoleName = TOMModelRole.Name;
                TOMModelRolePermission = TOMModelRole.ModelPermission.ToString();
                TOMModelRoleDescription = TOMModelRole.Description;
                TOMModelRoleMembers = "";
                foreach (TOM.ModelRoleMember TOMModelRoleMember in TOMModelRole.Members)
                {
                    TOMModelRoleMembers = TOMModelRoleMembers + TOMModelRoleMember.MemberName + Environment.NewLine;

                }

                foreach(TOM.TablePermission TOMTablePermission in TOMModelRole.TablePermissions )
                {
                    //Role Name
                    ExcelColNumber = 1;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMModelRoleName;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Role Permission
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMModelRolePermission;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);


                    //Role Description
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMModelRoleDescription;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Role Members
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMModelRoleMembers;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //Table
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMTablePermission.Table.Name; 
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //FilterExpression
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMTablePermission.FilterExpression; 
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    //MetadataPermission
                    ExcelColNumber++;
                    XLWorkSheet.Cells[ExcelSheetStartrow, ExcelColNumber] = TOMTablePermission.MetadataPermission.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, ExcelColNumber, -1);

                    ExcelSheetStartrow++;

                    

                }


            }
           
            
            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Roles Completed " + Environment.NewLine);
        }



        public void WriteHeaderCell(Worksheet XLWorkSheet, int row, int col, string headercaption)
        {



            Range CellRange;

            CellRange = (Microsoft.Office.Interop.Excel.Range)XLWorkSheet.Cells[row, col];


            CellRange.Value = headercaption;

            CellRange.Interior.Color = System.Drawing.Color.CornflowerBlue;
            CellRange.Font.Color = System.Drawing.Color.White;


            CellRange.Font.Bold = true;

            CellRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            CellRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            CellRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            CellRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;



            CellRange.EntireColumn.AutoFit();

            






        }

        public void WriteDataCell(Worksheet XLWorkSheet, int row, int col, string CellValue)
        {

            Range CellRange;

            CellRange = (Microsoft.Office.Interop.Excel.Range)XLWorkSheet.Cells[row, col];

            CellRange.Value = CellValue;

            CellRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            CellRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            CellRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            CellRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            CellRange.EntireColumn.AutoFit();

            

        }

        public void FormatCell(Worksheet XLWorkSheet, int row, int col, int CellWidth)
        {
            /*
            //XLWorkSheet.Cells[row, col] = CellValue;

            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            XLWorkSheet.Cells[row, col].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            if (CellWidth < 0)
            {
                XLWorkSheet.Cells[row, col].Entirecolumn.Autofit();
            }
            */

        }

        public void FormatSheet(Worksheet XLWorkSheet)
        {

            //XLWorkSheet.Cells[row, col] = CellValue;
            /*
            Range theRange = (Range) XLWorkSheet.UsedRange;
                theRange.Select();
            theRange.Cells.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            theRange.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            theRange.Columns.AutoFit();
            theRange.RowHeight = 50;
            XLApp.ActiveWindow.Zoom = 80;
            */

            Range theRange = (Range)XLWorkSheet.UsedRange;

            //theRange.Cells.BorderAround2(XlLineStyle.xlContinuous);

            theRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            theRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;


            /*
            theRange.Cells.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            
            SelectionRange Se =  theRange.Select();
           //Selection.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            theRange.Cells.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
             * */
            theRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            theRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            theRange.Columns.AutoFit();
            theRange.Rows.AutoFit();
            //theRange.RowHeight = 50;
            XLApp.ActiveWindow.Zoom = 80;



            //rn =XLWorkSheet.Range["A1"].Select();
            /*
            XLApp.Range[]
            Range("A1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Range(Selection, Selection.End(xlToRight)).Select
             */

        }

      


    }
}



