using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;

using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;


namespace BISMDocumenterAMO
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

        BISMDocumenterLibrary.ProgressWritter PX = new BISMDocumenterLibrary.ProgressWritter();
        //string  txtProgress;

        int ExcelSheetStartrow;

        public void GenerateDocument(String ServerName, String DBName, String CubeName, String DocumentPath,String FileName, System.Windows.Forms.TextBox progressTextBox,Boolean OpenXl)
        {
            try
            {
                
                String ConnStr;
                OLAPServerName = ServerName;
                // txtProgress= "";

                ConnStr = "Provider=MSOLAP;Data Source=" + OLAPServerName + ";";
                //Initial Catalog=Adventure Works DW 2008R2;"; 
                OLAPServer = new AMO.Server();
                OLAPServer.Connect(ConnStr);

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
                
                


                if (!System.IO.Directory.Exists(DocumentPath))
                {
                    System.IO.Directory.CreateDirectory(DocumentPath);
                }

                OutputPath = DocumentPath;


                OLAPDBName = DBName;
                OLAPCubeName = CubeName;

                OLAPDatabase = OLAPServer.Databases[OLAPDBName.Trim()];
                TOMDb = TOMServer.Databases[OLAPDBName.Trim()];

                Progressstr = "Database Compatibility Level - " + TOMDb.CompatibilityLevel.ToString();
                PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = "AMO Extracting Metadata for " + ServerName + " - " + DBName + " - " + CubeName;
                PX.WriteProgress(Progressstr, progressTextBox);

                //if (CubeName.Trim() == "")
                //{
                FileName = FileName + ".xlsx";
                //}
                //else
                //{
                //    Filename = DBName + "-" + CubeName + ".xlsx";
                //}

                XLApp = new Microsoft.Office.Interop.Excel.Application();
                XLApp.Visible = false;

                
                XLApp.DisplayAlerts = false;
                XLWorkBook = XLApp.Workbooks.Add();

                XLWorkBook.SaveAs(OutputPath + "\\" + FileName);
                

                
                progressTextBox.AppendText(XLWorkBook.Sheets.Count.ToString() + Environment.NewLine);
                XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add();
                XLWorkSheet.Name = "Server";
                
                OLAPCube = OLAPDatabase.Cubes.FindByName(OLAPCubeName);

                    WriteHeaderCell(XLWorkSheet, 1, 1, "Server Name");
                    XLWorkSheet.Cells[1, 2] = ServerName;
                    FormatCell(XLWorkSheet, 1, 2, -1);

                    WriteHeaderCell(XLWorkSheet, 2, 1, "Database Name");
                    XLWorkSheet.Cells[2, 2] = OLAPDBName;
                    FormatCell(XLWorkSheet, 2, 2, -1);

                    WriteHeaderCell(XLWorkSheet, 3, 1, "Cube Name");
                    XLWorkSheet.Cells[3, 2] = OLAPCubeName;
                    FormatCell(XLWorkSheet, 3, 2, -1);
                
                     XLWorkBook.Save();

                    String ProgressStringStartTemplate = "Generating Documentation for <PlaceHolder>....";

                    Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Connections");
                    PX.WriteProgress(Progressstr, progressTextBox);

                    Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Connections");
                    PX.WriteProgress(Progressstr, progressTextBox);

                    AMOExtractConnections();
                    FormatSheet(XLWorkSheet);

                    Progressstr = Progressstr.Replace("....", " Completed");
                    PX.WriteProgress(Progressstr, progressTextBox);

                    Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Dimensions");
                    PX.WriteProgress(Progressstr, progressTextBox);

                    AMOExtractDimension();
                    FormatSheet(XLWorkSheet);

                    Progressstr = Progressstr.Replace("....", " Completed");
                    PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Dimension Attributes");
                PX.WriteProgress(Progressstr, progressTextBox);

                AMOExtractDimensionAttribute();
                    FormatSheet(XLWorkSheet);

                    Progressstr = Progressstr.Replace("....", " Completed");
                    PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Relationships");
                PX.WriteProgress(Progressstr, progressTextBox);


                AMOExtractRelationship();
                    FormatSheet(XLWorkSheet);

                    Progressstr = Progressstr.Replace("....", " Completed");
                    PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Hierarchies");
                PX.WriteProgress(Progressstr, progressTextBox);


                AMOExtractHierarchies();
                    FormatSheet(XLWorkSheet);

                    Progressstr = Progressstr.Replace("....", " Completed");
                    PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Measures");
                PX.WriteProgress(Progressstr, progressTextBox);


                AMOExtractMeasures();
                    FormatSheet(XLWorkSheet);

                    Progressstr = Progressstr.Replace("....", " Completed");
                    PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "KPIs");
                PX.WriteProgress(Progressstr, progressTextBox);


                AMOExtractKPIs();
                    FormatSheet(XLWorkSheet);

                    Progressstr = Progressstr.Replace("....", " Completed");
                    PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Partitions");
                PX.WriteProgress(Progressstr, progressTextBox);


                AMOExtractPartitions();
                    FormatSheet(XLWorkSheet);

                    Progressstr = Progressstr.Replace("....", " Completed");
                    PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Perspectives");
                PX.WriteProgress(Progressstr, progressTextBox);


                AMOExtractPerspectives();
                    FormatSheet(XLWorkSheet);

                    Progressstr = Progressstr.Replace("....", " Completed");
                    PX.WriteProgress(Progressstr, progressTextBox);

                Progressstr = ProgressStringStartTemplate.Replace("<PlaceHolder>", "Roles");
                PX.WriteProgress(Progressstr, progressTextBox);


                AMOExtractRole();
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

                XLWorkBook.Save();
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
                Progressstr = "--------------------------------------------------------------------------------------" + Environment.NewLine;
                Progressstr = Progressstr + "Error Occured" + Environment.NewLine;
                Progressstr = Progressstr + "--------------------------------------------------------------------------------------" + Environment.NewLine;
                Progressstr = Progressstr + errormsg;
                
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

        public void AMOExtractConnections()
        {
            // txtProgress.AppendText(Progressstr + " Connections Started " + Environment.NewLine);

            WriteHeaderCell(XLWorkSheet, 5, 1, "Connections");
            WriteHeaderCell(XLWorkSheet, 6, 1, "Connection Name");
            WriteHeaderCell(XLWorkSheet, 6, 2, "Connection String");
            WriteHeaderCell(XLWorkSheet, 6, 3, "Description");



            ExcelSheetStartrow = 7;

            foreach (AMO.DataSource OlapDS in OLAPDatabase.DataSources)
            {

                XLWorkSheet.Cells[ExcelSheetStartrow, 1] = OlapDS.Name;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                XLWorkSheet.Cells[ExcelSheetStartrow, 2] = OlapDS.ConnectionString;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                XLWorkSheet.Cells[ExcelSheetStartrow, 3] = OlapDS.Description;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);

                ExcelSheetStartrow++;

            }
            XLWorkBook.Save();
            // txtProgress.AppendText(Progressstr + " Connections Completed " + Environment.NewLine);
        }

        public void AMOExtractDimension()
        {
            
            ExcelSheetStartrow = 2;


            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Server"]);
            XLWorkSheet.Name = "Dimensions";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Dimensions"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "DimensionName");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Description");
            // WriteHeaderCell(XLWorkSheet, 1, 3, "Hidden");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Connection");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Source Friendly Name");
            WriteHeaderCell(XLWorkSheet, 1, 5, "Source Schema Name");
            WriteHeaderCell(XLWorkSheet, 1, 6, "Source Table Name");
            WriteHeaderCell(XLWorkSheet, 1, 7, "Source Description");
            WriteHeaderCell(XLWorkSheet, 1, 8, "Source Query");





            foreach (AMO.Dimension Dimension in OLAPDatabase.Dimensions)
            {

                XLWorkSheet.Cells[ExcelSheetStartrow, 1] = Dimension.Name;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                XLWorkSheet.Cells[ExcelSheetStartrow, 2] = Dimension.Description;
                //XLWorkSheet.Cells[ExcelSheetStartrow, 2].Wraptext = true;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                // XLWorkSheet.Cells[ExcelSheetStartrow, 3] = "";
                // FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                XLWorkSheet.Cells[ExcelSheetStartrow, 3] = Dimension.DataSource.Name;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);

                AMO.DataSourceView OLAPDataSourceView = OLAPDatabase.DataSourceViews.Find(Dimension.DataSourceView.ID);
                if (OLAPDataSourceView.Schema.Tables[Dimension.ID] != null) // FRM Cube Temp Fix
                {
                    XLWorkSheet.Cells[ExcelSheetStartrow, 4] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["FriendlyName"];
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 5] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["DbSchemaName"];
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 5, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 6] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["DbTableName"];
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 6, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 7] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["Description"];
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 7, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 8] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["QueryDefinition"];
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 8, -1);
                    //XLWorkSheet.Cells[ExcelSheetStartrow, 7] = OLAPDataSourceView.Schema.Tables[0];
                }




                ExcelSheetStartrow++;

            }
            XLWorkBook.Save();

            

        }

        public void AMOExtractDimensionAttribute()
        {
            // txtProgress.AppendText(Progressstr + " Dimension Attributes Started " + Environment.NewLine);
            string ColumnSource;
            ExcelSheetStartrow = 2;
            // XLWorkSheet = XLWorkBook.Sheets["DimensionAttributes"];
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Dimensions"]);
            XLWorkSheet.Name = "DimensionAttributes";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["DimensionAttributes"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Dimension Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Attribute Name");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Description");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Data Type");
            WriteHeaderCell(XLWorkSheet, 1, 5, "Length");
            WriteHeaderCell(XLWorkSheet, 1, 6, "Column Name");
            WriteHeaderCell(XLWorkSheet, 1, 7, "FriendlyColumnName");
            WriteHeaderCell(XLWorkSheet, 1, 8, "DbColumnName");
            WriteHeaderCell(XLWorkSheet, 1, 9, "Calculated Column");
            WriteHeaderCell(XLWorkSheet, 1, 10, "Formula");
            WriteHeaderCell(XLWorkSheet, 1, 11, "Visible");
            WriteHeaderCell(XLWorkSheet, 1, 12, "Sort by Column");


            foreach (AMO.Dimension Dimension in OLAPDatabase.Dimensions)
            {
                foreach (AMO.DimensionAttribute DimAttribute in Dimension.Attributes)
                {
                    if (DimAttribute.Name.ToUpper() != "ROWNUMBER" && DimAttribute.Name.ToUpper() != "__XL_RowNumber".ToUpper())
                    {


                        XLWorkSheet.Cells[ExcelSheetStartrow, 1] = Dimension.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 2] = DimAttribute.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);

                        string desc = DimAttribute.Description;

                        /* Logic to modify the error due to = as first char */
                        if (desc != null)
                        {
                            if (desc.IndexOf("=") == 0)
                            {
                                desc = desc.Substring(1);
                            }
                        }

                        XLWorkSheet.Cells[ExcelSheetStartrow, 3] = desc;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);

                        XLWorkSheet.Cells[ExcelSheetStartrow, 4] = DimAttribute.NameColumn;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 5] = DimAttribute.NameColumn.DataSize;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 5, -1);


                        ColumnSource = DimAttribute.NameColumn.Source.ToString().Replace(Dimension.ID + ".", "");



                        if (ColumnSource == "Microsoft.AnalysisServices.ExpressionBinding")
                        {

                            //MessageBox.Show(((Microsoft.AnalysisServices.ExpressionBinding)DimAttribute.NameColumn.Source).Expression.ToString());
                            XLWorkSheet.Cells[ExcelSheetStartrow, 6] = DimAttribute.Name;
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 6, -1);
                            XLWorkSheet.Cells[ExcelSheetStartrow, 7] = "";
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 7, -1);
                            XLWorkSheet.Cells[ExcelSheetStartrow, 8] = "";
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 8, -1);
                            XLWorkSheet.Cells[ExcelSheetStartrow, 9] = "Yes";
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 9, -1);
                            XLWorkSheet.Cells[ExcelSheetStartrow, 10] = ((Microsoft.AnalysisServices.ExpressionBinding)DimAttribute.NameColumn.Source).Expression.ToString();
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 10, -1);
                        }
                        else
                        {
                            XLWorkSheet.Cells[ExcelSheetStartrow, 6] = ColumnSource;
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 6, -1);
                            AMO.DataSourceView OLAPDataSourceView = OLAPDatabase.DataSourceViews.Find(Dimension.DataSourceView.ID);

                            if (OLAPDataSourceView.Schema.Tables[Dimension.ID] != null) // FRM Cube Temp Fix
                            {
                                XLWorkSheet.Cells[ExcelSheetStartrow, 4] = OLAPDataSourceView.Schema.Tables[Dimension.ID].Columns[ColumnSource].DataType.UnderlyingSystemType.ToString();
                                FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);
                                XLWorkSheet.Cells[ExcelSheetStartrow, 5] = OLAPDataSourceView.Schema.Tables[Dimension.ID].Columns[ColumnSource].MaxLength;
                                FormatCell(XLWorkSheet, ExcelSheetStartrow, 5, -1);
                                XLWorkSheet.Cells[ExcelSheetStartrow, 7] = OLAPDataSourceView.Schema.Tables[Dimension.ID].Columns[ColumnSource].ExtendedProperties["FriendlyName"];
                                FormatCell(XLWorkSheet, ExcelSheetStartrow, 7, -1);
                                XLWorkSheet.Cells[ExcelSheetStartrow, 8] = OLAPDataSourceView.Schema.Tables[Dimension.ID].Columns[ColumnSource].ExtendedProperties["DbColumnName"];
                                FormatCell(XLWorkSheet, ExcelSheetStartrow, 8, -1);
                            }
                            XLWorkSheet.Cells[ExcelSheetStartrow, 9] = "No";
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 9, -1);
                            XLWorkSheet.Cells[ExcelSheetStartrow, 10] = "";
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 10, -1);

                        }

                        XLWorkSheet.Cells[ExcelSheetStartrow, 11] = DimAttribute.AttributeHierarchyVisible;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 11, -1);
                        if (DimAttribute.OrderByAttribute != null)
                        {

                            XLWorkSheet.Cells[ExcelSheetStartrow, 12] = DimAttribute.OrderByAttribute.NameColumn.Source.ToString().Replace(Dimension.ID + ".", "");
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 12, -1);
                        }


                        ExcelSheetStartrow++;
                    }
                }
            }
            XLWorkBook.Save();

            

        }

        public void AMOExtractRelationship()
        {
            // txtProgress.AppendText(Progressstr + " Relationships Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["DimensionAttributes"]);
            XLWorkSheet.Name = "Relationships";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Relationships"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "From Dimension");
            WriteHeaderCell(XLWorkSheet, 1, 2, "From Attributes");
            WriteHeaderCell(XLWorkSheet, 1, 3, "From Multiplicity");
            WriteHeaderCell(XLWorkSheet, 1, 4, "To Dimension");
            WriteHeaderCell(XLWorkSheet, 1, 5, "To Attributes");
            WriteHeaderCell(XLWorkSheet, 1, 6, "To Multiplicity");



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
            XLWorkBook.Save();

            

        }

        public void AMOExtractHierarchies()
        {
            // txtProgress.AppendText(Progressstr + " Hierarchies Started " + Environment.NewLine);
            int Hierarchylvl = 0;
            ExcelSheetStartrow = 2;

            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Relationships"]);
            XLWorkSheet.Name = "Hierarchies";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Hierarchies"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Dimension Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Hierarchy Name");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Level");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Level Name");
            WriteHeaderCell(XLWorkSheet, 1, 5, "Level Attribute Name");


            foreach (AMO.Dimension Dimension in OLAPDatabase.Dimensions)
            {
                foreach (AMO.Hierarchy DimHierarchy in Dimension.Hierarchies)
                {
                    Hierarchylvl = 1;
                    foreach (AMO.Level DimHierarchyLevel in DimHierarchy.Levels)
                    {
                        XLWorkSheet.Cells[ExcelSheetStartrow, 1] = Dimension.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 2] = DimHierarchy.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 3] = Hierarchylvl;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 4] = DimHierarchyLevel.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 5] = DimHierarchyLevel.SourceAttribute.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 5, -1);
                        ExcelSheetStartrow++;
                        Hierarchylvl++;
                    }
                }
            }
            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Hierarchies Completed " + Environment.NewLine);

        }

        public void AMOExtractMeasures()
        {
            // txtProgress.AppendText(Progressstr + " Measures Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;


            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Hierarchies"]);
            XLWorkSheet.Name = "Measures";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Measures"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Measure Group Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Measure Name");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Measure Expression");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Measure Description");


            /*
            foreach(MeasureGroup MsrGroup in OLAPCube.MeasureGroups)
            {
                foreach(Measure Msr in MsrGroup.Measures)
                {
                    XLWorkSheet.Cells[ExcelSheetStartrow, 1] = MsrGroup.Name;
                    XLWorkSheet.Cells[ExcelSheetStartrow, 2] = Msr.Name;
                    XLWorkSheet.Cells[ExcelSheetStartrow, 3] = Msr.Description;
                    XLWorkSheet.Cells[ExcelSheetStartrow, 4] = Msr.MeasureExpression;
                    XLWorkSheet.Cells[ExcelSheetStartrow, 5] = Msr.Visible;
                    ExcelSheetStartrow++;

                }
            }
            XLWorkBook.Save();
             * */

            string MeasureScript;
            string MeasureName;
            string MeasureFormula;

            MeasureScript = "";

            foreach (AMO.MdxScript MDXScript in OLAPCube.MdxScripts)
            {
                foreach (AMO.Command MDXCommand in MDXScript.Commands)
                {
                    MeasureScript = MeasureScript + Environment.NewLine + MDXCommand.Text;
                }
            }

            // MeasureScript = MeasureScript.Replace(Environment.NewLine, "");

            String[] MeasureArray = MeasureScript.Split(new string[] { "\nCREATE" }, StringSplitOptions.RemoveEmptyEntries);
            

            foreach (AMO.CubeDimension MeasureDimension in OLAPCube.Dimensions)
            {
                for (int i = 0; i <= MeasureArray.LongLength - 1; i++)
                {
                    if (MeasureArray[i].IndexOf("MEASURE '" + MeasureDimension.Name + "'") > 0)
                    {

                        MeasureName = MeasureArray[i].Substring(MeasureArray[i].IndexOf("["), MeasureArray[i].IndexOf("]") - MeasureArray[i].IndexOf("[") + 1);

                        

                        if (MeasureName.IndexOf("[_") < 0)
                        {
                            XLWorkSheet.Cells[ExcelSheetStartrow, 1] = MeasureDimension.Name;  // Dimension Name
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                            XLWorkSheet.Cells[ExcelSheetStartrow, 2] = MeasureName;         //Measure Name
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                            // XLWorkSheet.Cells[ExcelSheetStartrow, 3] = "";   //Description
                            MeasureFormula = MeasureArray[i].Substring(MeasureArray[i].IndexOf("=") + 1, MeasureArray[i].Length - (MeasureArray[i].IndexOf("=") + 1));  // MeasureArray[i].Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries)[1];
                            MeasureFormula = MeasureFormula.Substring(0, MeasureFormula.IndexOf(";"));
                            XLWorkSheet.Cells[ExcelSheetStartrow, 3] = MeasureFormula; //Formula
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                            XLWorkSheet.Cells[ExcelSheetStartrow, 4] = MeasureDimension.Description; //Description
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);


                            //XLWorkSheet.Cells[ExcelSheetStartrow, 5] = "";  //Visibility
                            ExcelSheetStartrow++;
                        }
                    }

                }



            }

            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Measures Completed " + Environment.NewLine);

        }

        public void AMOExtractKPIs()
        {
            // txtProgress.AppendText(Progressstr + " KPIs Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;

            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Measures"]);
            XLWorkSheet.Name = "KPIs";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["KPIs"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Measure Group Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "KPI Name");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Goal");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Value");
            WriteHeaderCell(XLWorkSheet, 1, 5, "Trend Graphic");


            /*
            foreach (Kpi DimKPI in OLAPCube.Kpis)
            {

                XLWorkSheet.Cells[ExcelSheetStartrow, 1] = DimKPI.AssociatedMeasureGroup.ToString();
                XLWorkSheet.Cells[ExcelSheetStartrow, 2] = DimKPI.Name.ToString();
                XLWorkSheet.Cells[ExcelSheetStartrow, 3] = DimKPI.Goal.ToString();
                XLWorkSheet.Cells[ExcelSheetStartrow, 4] = DimKPI.Value.ToString();
                XLWorkSheet.Cells[ExcelSheetStartrow, 5] = DimKPI.Trend.ToString();
                XLWorkSheet.Cells[ExcelSheetStartrow, 5] = DimKPI.TrendGraphic.ToString();


                ExcelSheetStartrow++;

            }
             * */

            string MeasureScript;
            string KPIName;
            string KPIAssocitaedMsrGroup;
            string KPIGoal;
            String KPIStatus;
            string KPIGoalValue = "";
            string KPIStatusValue = "";
            String KPIStatusGraphic;
            //string MeasureFormula;

            MeasureScript = "";

            foreach (AMO.MdxScript MDXScript in OLAPCube.MdxScripts)
            {
                foreach (AMO.Command MDXCommand in MDXScript.Commands)
                {
                    MeasureScript = MeasureScript + Environment.NewLine + MDXCommand.Text;
                }
            }

            String[] MeasureArray = MeasureScript.Split(new string[] { "CREATE" }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i <= MeasureArray.LongLength - 1; i++)
            {
                if (MeasureArray[i].IndexOf("KPI") >= 0)
                {

                    KPIAssocitaedMsrGroup = MeasureArray[i].Split(new string[] { "ASSOCIATED_MEASURE_GROUP =" }, StringSplitOptions.RemoveEmptyEntries)[1];
                    KPIAssocitaedMsrGroup = KPIAssocitaedMsrGroup.Trim().Substring(0, KPIAssocitaedMsrGroup.Trim().IndexOf("'", 1));

                    KPIName = MeasureArray[i].Substring(MeasureArray[i].IndexOf("["), MeasureArray[i].IndexOf("]") - MeasureArray[i].IndexOf("[") + 1);

                    KPIGoal = MeasureArray[i].Split(new string[] { "GOAL = Measures." }, StringSplitOptions.RemoveEmptyEntries)[1];
                    KPIGoal = KPIGoal.Substring(1, KPIGoal.IndexOf("]") - 1);

                    KPIStatus = MeasureArray[i].Split(new string[] { "STATUS = Measures." }, StringSplitOptions.RemoveEmptyEntries)[1];
                    KPIStatus = KPIStatus.Substring(1, KPIStatus.IndexOf("]") - 1);

                    KPIStatusGraphic = MeasureArray[i].Split(new string[] { "STATUS_GRAPHIC = '" }, StringSplitOptions.RemoveEmptyEntries)[1];
                    KPIStatusGraphic = KPIStatusGraphic.Substring(0, KPIStatusGraphic.IndexOf("'") - 1);

                    for (int x = 0; x <= MeasureArray.LongLength - 1; x++)
                    {
                        if (MeasureArray[x].IndexOf("[" + KPIGoal + "]") > 0 && MeasureArray[x].IndexOf("KPI") < 0)
                        {
                            KPIGoalValue = MeasureArray[x].Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries)[1];
                        }
                        if (MeasureArray[x].IndexOf("[" + KPIStatus + "]") > 0 && MeasureArray[x].IndexOf("KPI") < 0)
                        {
                            KPIStatusValue = MeasureArray[x].Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries)[1];
                        }
                    }
                    XLWorkSheet.Cells[ExcelSheetStartrow, 1] = KPIAssocitaedMsrGroup;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 2] = KPIName;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 3] = KPIGoalValue;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 4] = KPIStatusValue;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 5] = KPIStatusGraphic;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 5, -1);
                    //XLWorkSheet.Cells[ExcelSheetStartrow, 6] = MeasureArray[i]; 

                    ExcelSheetStartrow++;
                }
            }
            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " KPIs Completed " + Environment.NewLine);

        }



        public void AMOExtractPartitions()
        {
            // txtProgress.AppendText(Progressstr + " Partitions Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;

            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["KPIs"]);
            XLWorkSheet.Name = "Partitions";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Partitions"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Measure Group Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Source Type");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Source");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Estimated Rows");
            WriteHeaderCell(XLWorkSheet, 1, 5, "Estimated Size");


            foreach (AMO.MeasureGroup MsrGroup in OLAPCube.MeasureGroups)
            {
                foreach (AMO.Partition MsrGroupPartition in MsrGroup.Partitions)
                {
                    XLWorkSheet.Cells[ExcelSheetStartrow, 1] = MsrGroupPartition.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                    // "Microsoft.AnalysisServices.QueryBinding"
                    if (MsrGroupPartition.Source.ToString() == "Microsoft.AnalysisServices.QueryBinding")
                    {
                        XLWorkSheet.Cells[ExcelSheetStartrow, 2] = "QueryBinding";
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 3] = ((Microsoft.AnalysisServices.QueryBinding)MsrGroupPartition.Source).QueryDefinition;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                    }

                    if (MsrGroupPartition.Source.ToString() == "Microsoft.AnalysisServices.TableBinding")
                    {
                        XLWorkSheet.Cells[ExcelSheetStartrow, 2] = "TableBinding";
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 3] = ((Microsoft.AnalysisServices.TableBinding)MsrGroupPartition.Source).DbSchemaName + "." + ((Microsoft.AnalysisServices.TableBinding)MsrGroupPartition.Source).DbTableName;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                    }
                    XLWorkSheet.Cells[ExcelSheetStartrow, 4] = MsrGroupPartition.EstimatedRows;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 5] = MsrGroupPartition.EstimatedSize;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 5, -1);
                    ExcelSheetStartrow++;
                }
            }

            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Partitions Completed " + Environment.NewLine);

        }

        public void AMOExtractPerspectives()
        {
            // txtProgress.AppendText(Progressstr + " Perspectives Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;

            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Partitions"]);
            XLWorkSheet.Name = "Perspectives";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Perspectives"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Perspective Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Dimension Name");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Attribute Name");


            //OLAPCube.Perspectives[0].MeasureGroups[0].Measures[0].Measure.
            foreach (AMO.Perspective CubePerspective in OLAPCube.Perspectives)
            {
                foreach (AMO.PerspectiveDimension CubePerspectiveDim in CubePerspective.Dimensions)
                {

                    foreach (AMO.PerspectiveAttribute CubePerspectiveAttribute in CubePerspectiveDim.Attributes)
                    {
                        XLWorkSheet.Cells[ExcelSheetStartrow, 1] = CubePerspective.Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 2] = OLAPCube.Dimensions.Find(CubePerspectiveDim.CubeDimensionID).Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 3] = OLAPCube.Dimensions.Find(CubePerspectiveDim.CubeDimensionID).Attributes.Find(CubePerspectiveAttribute.AttributeID).Attribute.Name.ToString().Replace(CubePerspectiveDim.CubeDimensionID + ".", "");
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                        ExcelSheetStartrow++;
                    }
                }


                /*
                foreach (MeasureGroup CubePerspectiveMeasureGroup in CubePerspective.MeasureGroups)
                {
                    foreach(Measure CubePerspectiveMeasure in CubePerspectiveMeasureGroup.Measures )
                    {
                        XLWorkSheet.Cells[ExcelSheetStartrow, 1] = CubePerspective.Name;
                        XLWorkSheet.Cells[ExcelSheetStartrow, 2] = "Measure";
                        XLWorkSheet.Cells[ExcelSheetStartrow, 3] = CubePerspectiveMeasureGroup.Name + "-" + CubePerspectiveMeasure.Name;
                        ExcelSheetStartrow++;

                    }
                }
                 * */

            }
            XLWorkBook.Save();

            // txtProgress.AppendText(Progressstr + " Perspectives Completed " + Environment.NewLine);
        }

        public void AMOExtractRole()
        {
            // txtProgress.AppendText(Progressstr + " Roles Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;

            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Perspectives"]);
            XLWorkSheet.Name = "Roles";
            XLWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XLWorkBook.Sheets["Roles"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Role Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Role Description");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Adminster");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Process");
            WriteHeaderCell(XLWorkSheet, 1, 5, "Read");
            WriteHeaderCell(XLWorkSheet, 1, 6, "Dimension");
            WriteHeaderCell(XLWorkSheet, 1, 7, "RowFilter");


            foreach (AMO.DatabasePermission dbPermission in OLAPDatabase.DatabasePermissions)
            {

                foreach (AMO.Dimension Dim in OLAPDatabase.Dimensions)
                {
                    XLWorkSheet.Cells[ExcelSheetStartrow, 1] = dbPermission.Role.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 2] = dbPermission.Role.Description;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 3] = dbPermission.Administer;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 4] = dbPermission.Process;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 5] = dbPermission.Read.ToString();
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 5, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 6] = Dim.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 6, -1);

                    if (Dim.DimensionPermissions.Count > 0)
                    {
                        if (Dim.DimensionPermissions[0].RoleID == dbPermission.RoleID)
                        {
                            XLWorkSheet.Cells[ExcelSheetStartrow, 7] = Dim.DimensionPermissions[0].AllowedRowsExpression;
                            FormatCell(XLWorkSheet, ExcelSheetStartrow, 7, -1);
                        }
                    }
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

        public void TOMExtractConnections()
        {
            // txtProgress.AppendText(Progressstr + " Connections Started " + Environment.NewLine);

            int ColNumber = 1;

            WriteHeaderCell(XLWorkSheet, 5, 1, "Connections");

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


            ExcelSheetStartrow = 7;

            

            TOM.StructuredDataSource TOMDs;

            TOM.ProviderDataSource TOMPDs;
            
            ColNumber = 1;
            for (int I = 0;I <= TOMDb.Model.DataSources.Count-1;I++)
            {



                TOMDs = (TOM.StructuredDataSource)TOMDb.Model.DataSources[I];

                ColNumber = 1;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDb.Model.DataSources[I].Name;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.ConnectionString;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.Description;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
                ColNumber++;
                XLWorkSheet.Cells[ExcelSheetStartrow, ColNumber] = TOMDs.ConnectionDetails.Address.Server;
                FormatCell(XLWorkSheet, ExcelSheetStartrow, ColNumber, -1);
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


    }
}

