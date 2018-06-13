using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.AnalysisServices;
//using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;


namespace SSASDocumentationTool_DescriptionEditor
{
    public partial class frmSSASDocumentationTool : Form
    {

        string AppPath;
        string OutputPath;

        Microsoft.Office.Interop.Excel.Application XLApp;
        Microsoft.Office.Interop.Excel.Workbook XLWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet XLWorkSheet;

        String OLAPServerName;
        String OLAPDBName;
        String OLAPCubeName;

        Server OLAPServer;
        Database OLAPDatabase;
        Cube OLAPCube;

        string Progressstr;

        int ExcelSheetStartrow;
        

        public frmSSASDocumentationTool()
        {
            InitializeComponent();
        }

        private void cmdConnect_Click(object sender, EventArgs e)
        {

            try
            {
                String ConnStr;
                OLAPServerName = txtServerName.Text;
                txtProgress.AppendText("");

                ConnStr = "Provider=MSOLAP;Data Source=" + OLAPServerName + ";";
                //Initial Catalog=Adventure Works DW 2008R2;"; 
                OLAPServer = new Server();

                
                

                    OLAPServer.Connect(ConnStr);
                
                Console.WriteLine("ServerName : " + OLAPServerName);

                cboDatabaseName.Items.Clear();
                foreach (Database OLAPDatabase in OLAPServer.Databases)
                {
                    ComboboxItem item = new ComboboxItem();
                    item.Text = OLAPDatabase.Name;
                    item.Value = OLAPDatabase.ID;

                    cboDatabaseName.Items.Add(item);

                }

            }
            catch(Exception err)
            {
                string errormsg = err.InnerException.ToString();
                txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                txtProgress.AppendText("Error Occured" + Environment.NewLine);
                txtProgress.AppendText(err.InnerException.ToString() + Environment.NewLine);
                MessageBox.Show(errormsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmdGenerateDocument_Click(object sender, EventArgs e)
        {
            try
            {
                progressGeneration.Value = 1;
                txtProgress.AppendText("");
                Progressstr = "Generation started....";
                txtProgress.AppendText(Progressstr );

                String Filename;
                
                AppPath = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);

                if (!System.IO.Directory.Exists(txtOutputPath.Text))
                {
                    System.IO.Directory.CreateDirectory(txtOutputPath.Text);
                }

                OutputPath = txtOutputPath.Text;
                        OLAPDBName = cboDatabaseName.SelectedItem.ToString();   //DBCubeName.Substring(0,DBCubeName.IndexOf("-")-1);
                        OLAPCubeName = cboCubeName.SelectedItem.ToString(); // DBCubeName.Substring(DBCubeName.IndexOf("-") + 1, DBCubeName.Length - (DBCubeName.IndexOf("-")+1));


                        OLAPDatabase = OLAPServer.Databases[OLAPDBName.Trim()];
                        OLAPCube = OLAPDatabase.Cubes.FindByName(OLAPCubeName);

                        Filename = txtFileName.Text + ".xlsx";

                        XLApp = new Microsoft.Office.Interop.Excel.Application();
                        XLApp.Visible = false;

                        XLApp.DisplayAlerts = false;
                        XLWorkBook = XLApp.Workbooks.Add();
                        XLWorkBook.SaveAs(OutputPath + "\\" + Filename);
                        XLWorkSheet = XLWorkBook.Sheets.Add();
                        XLWorkSheet.Name = "Server";
                        

                        WriteHeaderCell(XLWorkSheet, 1, 1, "Server Name");
                        XLWorkSheet.Cells[1, 2] = txtServerName.Text;
                        FormatCell(XLWorkSheet, 1, 2, -1);

                        WriteHeaderCell(XLWorkSheet, 2, 1, "Database Name");
                        XLWorkSheet.Cells[2, 2] = OLAPDBName;
                        FormatCell(XLWorkSheet, 2, 2, -1);

                        WriteHeaderCell(XLWorkSheet, 3, 1, "Cube Name");
                        XLWorkSheet.Cells[3, 2] = OLAPCubeName;
                        FormatCell(XLWorkSheet, 3, 2, -1);

                        Progressstr = "Extracting Metadata for " + OLAPDatabase + " - " + OLAPCube;
                        txtProgress.AppendText(Progressstr + " Started " + Environment.NewLine);

                        XLWorkBook.Save();

                        ExtractConnections();
                        FormatSheet(XLWorkSheet);
                        progressGeneration.Value = 1;
                        ExtractDimension();
                        FormatSheet(XLWorkSheet);
                        progressGeneration.Value = 2;
                        ExtractDimensionAttribute();
                        FormatSheet(XLWorkSheet);
                        progressGeneration.Value = 3;
                        ExtractRelationship();
                        FormatSheet(XLWorkSheet);
                        progressGeneration.Value = 4;
                        ExtractHierarchies();
                        FormatSheet(XLWorkSheet);
                        progressGeneration.Value = 5;
                        ExtractMeasures();
                        FormatSheet(XLWorkSheet);
                        progressGeneration.Value = 6;
                        ExtractKPIs();
                        FormatSheet(XLWorkSheet);
                        progressGeneration.Value = 7;
                        ExtractPartitions();
                        FormatSheet(XLWorkSheet);
                        progressGeneration.Value = 8;
                        ExtractPerspectives();
                        FormatSheet(XLWorkSheet);
                        progressGeneration.Value = 9;
                        ExtractRole();
                        FormatSheet(XLWorkSheet);


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


                        if (sheet1exists==true)
                        {
                            XLWorkBook.Sheets["Sheet1"].Delete();
                        }

                        if (sheet2exists == true)
                        {
                            XLWorkBook.Sheets["Sheet2"].Delete();
                        }

                        if (sheet3exists == true)
                        {
                            XLWorkBook.Sheets["Sheet3"].Delete();
                        }
                        XLWorkBook.Sheets["Server"].Activate();
                        

                        txtProgress.AppendText(Progressstr + " Completed " + Environment.NewLine);
              //      }
              //  }

                XLWorkBook.Save();
                XLWorkBook.Close(true);
                XLApp.Quit();
                progressGeneration.Value = 10;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLApp);

                MessageBox.Show("Generation Completed Succesfully", "SSAS Documentation Tool", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch(Exception err)
            {

                string errormsg = err.ToString();
                txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                txtProgress.AppendText("Error Occured" + Environment.NewLine);
                txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                txtProgress.AppendText(err.ToString() + Environment.NewLine);
                MessageBox.Show(errormsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

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

        public void ExtractConnections()
        {
            txtProgress.AppendText( Progressstr + " Connections Started " + Environment.NewLine);

            WriteHeaderCell(XLWorkSheet, 5, 1, "Connections");
            WriteHeaderCell(XLWorkSheet, 6, 1, "Connection Name");
            WriteHeaderCell(XLWorkSheet, 6, 2, "Connection String");
            WriteHeaderCell(XLWorkSheet, 6, 3, "Description");
            
            

            ExcelSheetStartrow = 7;

            foreach ( DataSource OlapDS in OLAPDatabase.DataSources)
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
            txtProgress.AppendText( Progressstr + " Connections Completed " + Environment.NewLine);
        }

        public void ExtractDimension()
        {
           // try
           // {
                txtProgress.AppendText(Progressstr + " Dimensions Started " + Environment.NewLine);
                ExcelSheetStartrow = 2;

                XLWorkSheet = XLWorkBook.Sheets.Add(Type.Missing,XLWorkBook.Sheets["Server"]);
                XLWorkSheet.Name = "Dimensions";
                XLWorkSheet = XLWorkBook.Sheets["Dimensions"];

                WriteHeaderCell(XLWorkSheet, 1, 1, "DimensionName");
                WriteHeaderCell(XLWorkSheet, 1, 2, "Description");
               // WriteHeaderCell(XLWorkSheet, 1, 3, "Hidden");
                WriteHeaderCell(XLWorkSheet, 1, 3, "Connection");
                WriteHeaderCell(XLWorkSheet, 1, 4, "Source Friendly Name");
                WriteHeaderCell(XLWorkSheet, 1, 5, "Source Schema Name");
                WriteHeaderCell(XLWorkSheet, 1, 6, "Source Table Name");
                WriteHeaderCell(XLWorkSheet, 1, 7, "Source Description");
                WriteHeaderCell(XLWorkSheet, 1, 8, "Source Query");

            								

                

                foreach (Dimension Dimension in OLAPDatabase.Dimensions)
                {
                    
                    XLWorkSheet.Cells[ExcelSheetStartrow, 1] = Dimension.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 2] = Dimension.Description;
                    XLWorkSheet.Cells[ExcelSheetStartrow, 2].Wraptext = true;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                   // XLWorkSheet.Cells[ExcelSheetStartrow, 3] = "";
                   // FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                    XLWorkSheet.Cells[ExcelSheetStartrow, 3] = Dimension.DataSource.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);

                    DataSourceView OLAPDataSourceView = OLAPDatabase.DataSourceViews.Find(Dimension.DataSourceView.ID);
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

                txtProgress.AppendText(Progressstr + " Dimensions Completed " + Environment.NewLine);
            /*
             }
            catch (Exception err)
            {

                string errormsg = err.ToString();
                txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                txtProgress.AppendText("Error Occured" + Environment.NewLine);
                txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                txtProgress.AppendText(errormsg + Environment.NewLine);

                throw (err);

            }
             * */

        }

        public void ExtractDimensionAttribute()
        {
            txtProgress.AppendText( Progressstr + " Dimension Attributes Started " + Environment.NewLine);
            string ColumnSource;
            ExcelSheetStartrow = 2;
            // XLWorkSheet = XLWorkBook.Sheets["DimensionAttributes"];
            XLWorkSheet = XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Dimensions"]);
            XLWorkSheet.Name = "DimensionAttributes";
            XLWorkSheet = XLWorkBook.Sheets["DimensionAttributes"];

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


            foreach (Dimension Dimension in OLAPDatabase.Dimensions)
            {
                foreach (DimensionAttribute DimAttribute in Dimension.Attributes)
                {
                    if (DimAttribute.Name.ToUpper() != "ROWNUMBER" && DimAttribute.Name.ToUpper() != "__XL_RowNumber".ToUpper())
                    {

                        //if (Dimension.Name == "VAR_AH")
                        // {
                        //     MessageBox.Show("Test");
                        // }

                        //MessageBox.Show(Dimension.Name);

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
                        XLWorkSheet.Cells[ExcelSheetStartrow, 4] = DimAttribute.NameColumn.DataType.ToString();
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 4, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 5] = DimAttribute.NameColumn.DataSize;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 5, -1);
                        

                        ColumnSource = DimAttribute.NameColumn.Source.ToString().Replace(Dimension.ID + ".","");
                        
                        

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
                            DataSourceView OLAPDataSourceView = OLAPDatabase.DataSourceViews.Find(Dimension.DataSourceView.ID);

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

            txtProgress.AppendText( Progressstr + " Dimension Attributes Completed " + Environment.NewLine);

        }

        public void ExtractRelationship()
        {
            txtProgress.AppendText(Progressstr + " Relationships Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;
            XLWorkSheet = XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["DimensionAttributes"]);
            XLWorkSheet.Name = "Relationships";
            XLWorkSheet = XLWorkBook.Sheets["Relationships"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "From Dimension");
            WriteHeaderCell(XLWorkSheet, 1, 2, "From Attributes");
            WriteHeaderCell(XLWorkSheet, 1, 3, "From Multiplicity");
            WriteHeaderCell(XLWorkSheet, 1, 4, "To Dimension");
            WriteHeaderCell(XLWorkSheet, 1, 5, "To Attributes");
            WriteHeaderCell(XLWorkSheet, 1, 6, "To Multiplicity");



            foreach (Dimension RelDimension in OLAPDatabase.Dimensions)
            {
                foreach (Relationship DimRelationship in RelDimension.Relationships)
                {
                    XLWorkSheet.Cells[ExcelSheetStartrow, 1] = RelDimension.Name;
                    FormatCell(XLWorkSheet, ExcelSheetStartrow, 1, -1);
                    foreach (RelationshipEndAttribute FromRelAttribute in DimRelationship.FromRelationshipEnd.Attributes)
                    {
                        XLWorkSheet.Cells[ExcelSheetStartrow, 2] = RelDimension.Attributes[FromRelAttribute.AttributeID.ToString()].Name;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 2, -1);
                        XLWorkSheet.Cells[ExcelSheetStartrow, 3] = DimRelationship.FromRelationshipEnd.Multiplicity;
                        FormatCell(XLWorkSheet, ExcelSheetStartrow, 3, -1);
                    }
                    foreach (RelationshipEndAttribute ToRelAttribute in DimRelationship.ToRelationshipEnd.Attributes)
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

            txtProgress.AppendText(Progressstr + " Relationships Completed " + Environment.NewLine);

        }

        public void ExtractHierarchies()
        {
            txtProgress.AppendText(Progressstr + " Hierarchies Started " + Environment.NewLine);
            int Hierarchylvl = 0;
            ExcelSheetStartrow = 2;

            XLWorkSheet = XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Relationships"]);
            XLWorkSheet.Name = "Hierarchies";
            XLWorkSheet = XLWorkBook.Sheets["Hierarchies"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Dimension Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Hierarchy Name");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Level");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Level Name");
            WriteHeaderCell(XLWorkSheet, 1, 5, "Level Attribute Name");


            foreach (Dimension Dimension in OLAPDatabase.Dimensions)
            {
                foreach (Hierarchy DimHierarchy in Dimension.Hierarchies)
                {
                    Hierarchylvl = 1;
                    foreach (Level DimHierarchyLevel in DimHierarchy.Levels)
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

            txtProgress.AppendText(Progressstr + " Hierarchies Completed " + Environment.NewLine);

        }

        public void ExtractMeasures()
        {
            txtProgress.AppendText( Progressstr + " Measures Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;


            XLWorkSheet = XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Hierarchies"]);
            XLWorkSheet.Name = "Measures";
            XLWorkSheet = XLWorkBook.Sheets["Measures"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Measure Group Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Measure Name");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Measure Expression");


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

            foreach(MdxScript MDXScript in OLAPCube.MdxScripts)
            {
                foreach(Command MDXCommand in MDXScript.Commands)
                {
                    MeasureScript = MeasureScript + Environment.NewLine + MDXCommand.Text;
                }
            }

            // MeasureScript = MeasureScript.Replace(Environment.NewLine, "");

           String[] MeasureArray =   MeasureScript.Split( new string[]{ "\nCREATE" }, StringSplitOptions.RemoveEmptyEntries);


            foreach (CubeDimension MeasureDimension in OLAPCube.Dimensions)
            {
                for (int i = 0; i <= MeasureArray.LongLength-1;i++ )
                {
                    if (MeasureArray[i].IndexOf("MEASURE '" + MeasureDimension.Name + "'") > 0)
                    {
                        
                        MeasureName = MeasureArray[i].Substring(MeasureArray[i].IndexOf("["), MeasureArray[i].IndexOf("]") - MeasureArray[i].IndexOf("[")+1);

                        

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
                            //XLWorkSheet.Cells[ExcelSheetStartrow, 5] = "";  //Visibility
                            ExcelSheetStartrow++;
                        }
                    }

                }

              

            }
             
            XLWorkBook.Save();

            txtProgress.AppendText( Progressstr + " Measures Completed " + Environment.NewLine);

        }

        public void ExtractKPIs()
        {
            txtProgress.AppendText( Progressstr + " KPIs Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;

            XLWorkSheet = XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Measures"]);
            XLWorkSheet.Name = "KPIs";
            XLWorkSheet = XLWorkBook.Sheets["KPIs"];

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

            foreach (MdxScript MDXScript in OLAPCube.MdxScripts)
            {
                foreach (Command MDXCommand in MDXScript.Commands)
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
                            KPIAssocitaedMsrGroup = KPIAssocitaedMsrGroup.Trim().Substring(0, KPIAssocitaedMsrGroup.Trim().IndexOf("'",1));
                        
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

            txtProgress.AppendText( Progressstr + " KPIs Completed " + Environment.NewLine);

        }

        

        public void ExtractPartitions()
        {
            txtProgress.AppendText( Progressstr + " Partitions Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;

            XLWorkSheet = XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["KPIs"]);
            XLWorkSheet.Name = "Partitions";
            XLWorkSheet = XLWorkBook.Sheets["Partitions"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Measure Group Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Source Type");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Source");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Estimated Rows");
            WriteHeaderCell(XLWorkSheet, 1, 5, "Estimated Size");

            
            foreach(MeasureGroup MsrGroup in OLAPCube.MeasureGroups)
            {
                foreach(Partition MsrGroupPartition in MsrGroup.Partitions)
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

            txtProgress.AppendText( Progressstr + " Partitions Completed " + Environment.NewLine);

        }

        public void ExtractPerspectives()
        {
            txtProgress.AppendText( Progressstr + " Perspectives Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;

            XLWorkSheet = XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Partitions"]);
            XLWorkSheet.Name = "Perspectives";
            XLWorkSheet = XLWorkBook.Sheets["Perspectives"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Perspective Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Dimension Name");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Attribute Name");

            
            //OLAPCube.Perspectives[0].MeasureGroups[0].Measures[0].Measure.
            foreach(Perspective CubePerspective in OLAPCube.Perspectives)
            {
                foreach (PerspectiveDimension CubePerspectiveDim in CubePerspective.Dimensions)
                {
                    
                    foreach (PerspectiveAttribute CubePerspectiveAttribute in CubePerspectiveDim.Attributes)
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

            txtProgress.AppendText( Progressstr + " Perspectives Completed " + Environment.NewLine);
        }

        public void ExtractRole()
        {
            txtProgress.AppendText( Progressstr + " Roles Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;

            XLWorkSheet = XLWorkBook.Sheets.Add(Type.Missing, XLWorkBook.Sheets["Perspectives"]);
            XLWorkSheet.Name = "Roles";
            XLWorkSheet = XLWorkBook.Sheets["Roles"];

            WriteHeaderCell(XLWorkSheet, 1, 1, "Role Name");
            WriteHeaderCell(XLWorkSheet, 1, 2, "Role Description");
            WriteHeaderCell(XLWorkSheet, 1, 3, "Adminster");
            WriteHeaderCell(XLWorkSheet, 1, 4, "Process");
            WriteHeaderCell(XLWorkSheet, 1, 5, "Read");
            WriteHeaderCell(XLWorkSheet, 1, 6, "Dimension");
            WriteHeaderCell(XLWorkSheet, 1, 7, "RowFilter");


            foreach(DatabasePermission dbPermission in OLAPDatabase.DatabasePermissions)
            {

                foreach(Dimension Dim in OLAPDatabase.Dimensions)
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

            txtProgress.AppendText( Progressstr + " Roles Completed " + Environment.NewLine);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
           txtOutputPath.Text =  System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\Output";
        }

        private void cboDatabaseName_SelectedIndexChanged(object sender, EventArgs e)
        {

            cboCubeName.Items.Clear();
            cboCubeName.SelectedItem = null;

            //OLAPDatabase = OLAPServer.Databases[cboDatabaseName.SelectedItem.ToString()];
            OLAPDatabase = OLAPServer.Databases[(cboDatabaseName.SelectedItem as ComboboxItem).Value.ToString()];
            
            foreach (Cube OLAPCube in OLAPDatabase.Cubes)
            {

                cboCubeName.Items.Add(OLAPCube.Name);

            }

            if (cboCubeName.Items.Count == 1)
            {
                cboCubeName.SelectedItem = cboCubeName.Items[0];
            }

        }

        private void cboCubeName_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtFileName.Text = cboDatabaseName.SelectedItem.ToString() + "-" + cboCubeName.SelectedItem.ToString();
        }

        public void WriteHeaderCell(Worksheet XLWorkSheet,int row,int col,string headercaption)
        {

            XLWorkSheet.Cells[row, col] = headercaption;
            XLWorkSheet.Cells[row, col].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);
            XLWorkSheet.Cells[row, col].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            XLWorkSheet.Cells[row, col].Font.Bold = true;

            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            XLWorkSheet.Cells[row, col].Entirecolumn.Autofit();

        }

        public void WriteDataCell(Worksheet XLWorkSheet, int row, int col, string CellValue)
        {

            XLWorkSheet.Cells[row, col] = CellValue;
           
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            XLWorkSheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            XLWorkSheet.Cells[row, col].Entirecolumn.Autofit();

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
            theRange.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            theRange.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
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


        public class ComboboxItem
        {
            public string Text { get; set; }
            public object Value { get; set; }

            public override string ToString()
            {
                return Text;
            }
        }

    }
}
