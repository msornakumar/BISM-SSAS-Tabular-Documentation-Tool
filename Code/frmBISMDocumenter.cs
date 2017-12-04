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
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace BISMDocumentor_DescriptionEditor
{
    public partial class frmBISMDocumentor : Form
    {
        
        string AppPath;
        string TemplatePath;
        string OutputPath;

        Microsoft.Office.Interop.Excel.Application ExcelApp;
        Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;

        String OLAPServerName;
        String OLAPDBName;
        String OLAPCubeName;

        Microsoft.AnalysisServices.Server OLAPServer;
        Database OLAPDatabase;
        Cube OLAPCube;

        string Progressstr;

        int ExcelSheetStartrow;
        

        public frmBISMDocumentor()
        {
            InitializeComponent();
        }

        private void cmdConnect_Click(object sender, EventArgs e)
        {

            try
            {
                String ConnStr;




                OLAPServerName = txtServerName.Text;


                clstDBCubeName.Items.Clear();

                txtProgress.AppendText("");

                ConnStr = "Provider=MSOLAP;Data Source=" + OLAPServerName + ";";
                //Initial Catalog=Adventure Works DW 2008R2;"; 

                OLAPServer = new Microsoft.AnalysisServices.Server();
                OLAPServer.Connect(ConnStr);

                Console.WriteLine("ServerName : " + OLAPServerName);

                // Database 
                foreach (Database OLAPDatabase in OLAPServer.Databases)
                {
                    // Cube 
                    foreach (Cube OLAPCube in OLAPDatabase.Cubes)
                    {

                        clstDBCubeName.Items.Add(OLAPDatabase.Name + " * " + OLAPCube.Name);

                    }
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
                txtProgress.AppendText("");

                String Filename;
                String DBCubeName;

                AppPath = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);

                if (!System.IO.Directory.Exists(txtOutputPath.Text))
                {

                    System.IO.Directory.CreateDirectory(txtOutputPath.Text);
                }

                OutputPath = txtOutputPath.Text;
                TemplatePath = AppPath + "\\Template\\BIDocumentTemplate.xlsx";



                for (int selindex = 0; selindex <= clstDBCubeName.Items.Count - 1; selindex++)
                {
                    if (clstDBCubeName.GetItemCheckState(selindex) == CheckState.Checked)
                    {
                        DBCubeName = clstDBCubeName.Items[selindex].ToString();
                        OLAPDBName = DBCubeName.Split('*')[0].Trim();   //DBCubeName.Substring(0,DBCubeName.IndexOf("-")-1);
                        OLAPCubeName = DBCubeName.Split('*')[1].Trim(); // DBCubeName.Substring(DBCubeName.IndexOf("-") + 1, DBCubeName.Length - (DBCubeName.IndexOf("-")+1));

                        OLAPDatabase = OLAPServer.Databases.FindByName(OLAPDBName.Trim());
                        OLAPCube = OLAPDatabase.Cubes.FindByName(OLAPCubeName);

                        Filename = DBCubeName.Replace('*', '-') + ".xlsx";
                        File.Copy(TemplatePath, OutputPath + "\\" + Filename, true);

                        ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                        ExcelApp.Visible = false;
                        ExcelWorkBook = ExcelApp.Workbooks.Open(OutputPath + "\\" + Filename);
                        ExcelWorkSheet = ExcelWorkBook.Sheets["Server"];

                        ExcelWorkSheet.Cells[1, 2] = txtServerName.Text;
                        ExcelWorkSheet.Cells[2, 2] = OLAPDBName;
                        ExcelWorkSheet.Cells[3, 2] = OLAPCubeName;

                        Progressstr = "Getting Metadata for " + OLAPDatabase + " - " + OLAPCube;
                        txtProgress.AppendText(Progressstr + " Started " + Environment.NewLine);

                        ExcelWorkBook.Save();

                        WriteConnections();
                        WriteDimension();
                        WriteDimensionAttribute();
                        WriteRelationship();
                        WriteHierarchies();
                        WriteMeasures();
                        WriteKPIs();
                        WritePartitions();
                        WritePerspectives();
                        WriteRole();

                        txtProgress.AppendText(Progressstr + " Completed " + Environment.NewLine);
                    }
                }

                ExcelWorkBook.Save();
                ExcelWorkBook.Close(true);
                ExcelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);

                MessageBox.Show("Excel Generation Completed", "BISM SSASTabularDocumetor", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch(Exception err)
            {

                string errormsg = err.ToString();
                txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                txtProgress.AppendText("Error Occured" + Environment.NewLine);
                txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                txtProgress.AppendText(err.ToString() + Environment.NewLine);
                MessageBox.Show(errormsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                ExcelWorkBook.Save();
                ExcelWorkBook.Close(true);
                ExcelApp.Quit();
                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);


            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                GC.Collect();
            }

        } 

        public void WriteConnections()
        {
            txtProgress.AppendText( Progressstr + " Connections Started " + Environment.NewLine);
            ExcelSheetStartrow = 7;

            foreach ( DataSource OlapDS in OLAPDatabase.DataSources)
            {
                ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = OlapDS.Name;
                ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = OlapDS.ConnectionString;
                ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = OlapDS.Description;
                ExcelSheetStartrow++;

            }
            ExcelWorkBook.Save();
            txtProgress.AppendText( Progressstr + " Connections Completed " + Environment.NewLine);
        }

        public void WriteDimension()
        {
           // try
           // {
                txtProgress.AppendText(Progressstr + " Dimensions Started " + Environment.NewLine);
                ExcelSheetStartrow = 2;
                ExcelWorkSheet = ExcelWorkBook.Sheets["Dimensions"];
                foreach (Dimension Dimension in OLAPDatabase.Dimensions)
                {

                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = Dimension.Name;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = Dimension.Description;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = "True";
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = Dimension.DataSource.Name;

                    DataSourceView OLAPDataSourceView = OLAPDatabase.DataSourceViews.Find(Dimension.DataSourceView.ID);
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["FriendlyName"];
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 6] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["DbSchemaName"];
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 7] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["DbTableName"];
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 8] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["Description"];
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 9] = OLAPDataSourceView.Schema.Tables[Dimension.ID].ExtendedProperties["QueryDefinition"];
                    //ExcelWorkSheet.Cells[ExcelSheetStartrow, 7] = OLAPDataSourceView.Schema.Tables[0];


                    ExcelSheetStartrow++;

                }
                ExcelWorkBook.Save();

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

        public void WriteDimensionAttribute()
        {
            txtProgress.AppendText( Progressstr + " Dimension Attributes Started " + Environment.NewLine);
            string ColumnSource;
            ExcelSheetStartrow = 2;
            ExcelWorkSheet = ExcelWorkBook.Sheets["DimensionAttributes"];
            foreach (Dimension Dimension in OLAPDatabase.Dimensions)
            {
                foreach (DimensionAttribute DimAttribute in Dimension.Attributes)
                {
                    if (DimAttribute.Name.ToUpper() != "ROWNUMBER")
                    {
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = Dimension.Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = DimAttribute.Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = DimAttribute.Description;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = DimAttribute.NameColumn.DataType.ToString();
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = DimAttribute.NameColumn.DataSize;
                       
                        

                        ColumnSource = DimAttribute.NameColumn.Source.ToString().Replace(Dimension.ID + ".","");
                        
                        

                        if (ColumnSource == "Microsoft.AnalysisServices.ExpressionBinding")
                        {

                            //MessageBox.Show(((Microsoft.AnalysisServices.ExpressionBinding)DimAttribute.NameColumn.Source).Expression.ToString());
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 6] = DimAttribute.Name;
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 7] = "";
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 8] = "";
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 9] = "Yes";
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 10] = ((Microsoft.AnalysisServices.ExpressionBinding)DimAttribute.NameColumn.Source).Expression.ToString();
                        }
                        else
                        {
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 6] = ColumnSource;

                            DataSourceView OLAPDataSourceView = OLAPDatabase.DataSourceViews.Find(Dimension.DataSourceView.ID);
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = OLAPDataSourceView.Schema.Tables[Dimension.ID].Columns[ColumnSource].DataType.UnderlyingSystemType.ToString();
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = OLAPDataSourceView.Schema.Tables[Dimension.ID].Columns[ColumnSource].MaxLength;
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 7] = OLAPDataSourceView.Schema.Tables[Dimension.ID].Columns[ColumnSource].ExtendedProperties["FriendlyName"];
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 8] = OLAPDataSourceView.Schema.Tables[Dimension.ID].Columns[ColumnSource].ExtendedProperties["DbColumnName"];
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 9] = "No";
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 10] = "";
                            
                        }

                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 11] = DimAttribute.AttributeHierarchyVisible;
                        
                        if (DimAttribute.OrderByAttribute != null)
                        {

                             ExcelWorkSheet.Cells[ExcelSheetStartrow, 12] = DimAttribute.OrderByAttribute.NameColumn.Source.ToString().Replace(Dimension.ID + ".", "");
                            
                        }


                        ExcelSheetStartrow++;
                    }
                }
            }
            ExcelWorkBook.Save();

            txtProgress.AppendText( Progressstr + " Dimension Attributes Completed " + Environment.NewLine);

        }

        public void WriteMeasures()
        {
            txtProgress.AppendText( Progressstr + " Measures Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;
            ExcelWorkSheet = ExcelWorkBook.Sheets["Measures"];

            /*
            foreach(MeasureGroup MsrGroup in OLAPCube.MeasureGroups)
            {
                foreach(Measure Msr in MsrGroup.Measures)
                {
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = MsrGroup.Name;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = Msr.Name;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = Msr.Description;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = Msr.MeasureExpression;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = Msr.Visible;
                    ExcelSheetStartrow++;

                }
            }
            ExcelWorkBook.Save();
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

            String[] MeasureArray =   MeasureScript.Split( new string[]{"CREATE"}, StringSplitOptions.RemoveEmptyEntries);


            foreach (CubeDimension MeasureDimension in OLAPCube.Dimensions)
            {
                for (int i = 0; i <= MeasureArray.LongLength-1;i++ )
                {
                    if (MeasureArray[i].IndexOf("MEASURE '" + MeasureDimension.Name + "'") > 0)
                    {
                        
                        MeasureName = MeasureArray[i].Substring(MeasureArray[i].IndexOf("["), MeasureArray[i].IndexOf("]") - MeasureArray[i].IndexOf("[")+1);

                        if (MeasureName.IndexOf("[_") < 0)
                        {
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = MeasureDimension.Name;  // Dimension Name
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = MeasureName;         //Measure Name
                           // ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = "";   //Description
                            MeasureFormula = MeasureArray[i].Substring(MeasureArray[i].IndexOf("=") + 1, MeasureArray[i].Length - (MeasureArray[i].IndexOf("=") + 1));  // MeasureArray[i].Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries)[1];
                            MeasureFormula = MeasureFormula.Substring(0, MeasureFormula.IndexOf(";"));
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = MeasureFormula; //Formula
                            //ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = "";  //Visibility
                            ExcelSheetStartrow++;
                        }
                    }

                }

              

            }
             
            ExcelWorkBook.Save();

            txtProgress.AppendText( Progressstr + " Measures Completed " + Environment.NewLine);

        }

        public void WriteKPIs()
        {
            txtProgress.AppendText( Progressstr + " KPIs Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;
            ExcelWorkSheet = ExcelWorkBook.Sheets["KPIs"];
            /*
            foreach (Kpi DimKPI in OLAPCube.Kpis)
            {

                ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = DimKPI.AssociatedMeasureGroup.ToString();
                ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = DimKPI.Name.ToString();
                ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = DimKPI.Goal.ToString();
                ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = DimKPI.Value.ToString();
                ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = DimKPI.Trend.ToString();
                ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = DimKPI.TrendGraphic.ToString();


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
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = KPIAssocitaedMsrGroup;  
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = KPIName;
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = KPIGoalValue;
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = KPIStatusValue; 
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = KPIStatusGraphic;
                            //ExcelWorkSheet.Cells[ExcelSheetStartrow, 6] = MeasureArray[i]; 
                            
                            ExcelSheetStartrow++;
                        }
                }
            ExcelWorkBook.Save();

            txtProgress.AppendText( Progressstr + " KPIs Completed " + Environment.NewLine);

        }

        public void WriteRelationship()
        {
            txtProgress.AppendText( Progressstr + " Relationships Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;
            ExcelWorkSheet = ExcelWorkBook.Sheets["Relationships"];
            foreach (Dimension RelDimension in OLAPDatabase.Dimensions)
            {
                foreach (Relationship DimRelationship in RelDimension.Relationships)
                {
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = RelDimension.Name;
                    foreach (RelationshipEndAttribute FromRelAttribute in DimRelationship.FromRelationshipEnd.Attributes)
                    {
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = RelDimension.Attributes[FromRelAttribute.AttributeID.ToString()].Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = DimRelationship.FromRelationshipEnd.Multiplicity;
                    }
                    foreach (RelationshipEndAttribute ToRelAttribute in DimRelationship.ToRelationshipEnd.Attributes)
                    {
                        
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = OLAPCube.Dimensions.Find( DimRelationship.ToRelationshipEnd.DimensionID).Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = OLAPCube.Dimensions.Find(DimRelationship.ToRelationshipEnd.DimensionID).Attributes.Find(ToRelAttribute.AttributeID);
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 6] = DimRelationship.ToRelationshipEnd.Multiplicity;
                    }

                    ExcelSheetStartrow++;
                }
                
                

            }
            ExcelWorkBook.Save();

            txtProgress.AppendText( Progressstr + " Relationships Completed " + Environment.NewLine);

        }

        public void WriteHierarchies()
        {
            txtProgress.AppendText( Progressstr + " Hierarchies Started " + Environment.NewLine);
            int Hierarchylvl = 0;
            ExcelSheetStartrow = 2;
            ExcelWorkSheet = ExcelWorkBook.Sheets["Hierarchies"];
            foreach (Dimension Dimension in OLAPDatabase.Dimensions)
            {
                foreach(Hierarchy DimHierarchy in Dimension.Hierarchies)
                {
                    Hierarchylvl = 1;
                    foreach(Level DimHierarchyLevel in DimHierarchy.Levels )
                    {
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = Dimension.Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = DimHierarchy.Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = Hierarchylvl;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = DimHierarchyLevel.Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = DimHierarchyLevel.SourceAttribute.Name;
                        ExcelSheetStartrow++;
                        Hierarchylvl++;
                    }
                }
            }
            ExcelWorkBook.Save();

            txtProgress.AppendText( Progressstr + " Hierarchies Completed " + Environment.NewLine);

        }

        public void WritePartitions()
        {
            txtProgress.AppendText( Progressstr + " Partitions Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;
            ExcelWorkSheet = ExcelWorkBook.Sheets["Partitions"];

            foreach(MeasureGroup MsrGroup in OLAPCube.MeasureGroups)
            {
                foreach(Partition MsrGroupPartition in MsrGroup.Partitions)
                {
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = MsrGroupPartition.Name;
                    // "Microsoft.AnalysisServices.QueryBinding"
                    if (MsrGroupPartition.Source.ToString() == "Microsoft.AnalysisServices.QueryBinding")
                    {
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = "QueryBinding";
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = ((Microsoft.AnalysisServices.QueryBinding)MsrGroupPartition.Source).QueryDefinition;
                    }

                    if (MsrGroupPartition.Source.ToString() == "Microsoft.AnalysisServices.TableBinding")
                    {
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = "TableBinding";
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = ((Microsoft.AnalysisServices.TableBinding)MsrGroupPartition.Source).DbSchemaName + "." + ((Microsoft.AnalysisServices.TableBinding)MsrGroupPartition.Source).DbTableName;
                    }
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = MsrGroupPartition.EstimatedRows;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = MsrGroupPartition.EstimatedSize;
                    ExcelSheetStartrow++;
                }
            }

            ExcelWorkBook.Save();

            txtProgress.AppendText( Progressstr + " Partitions Completed " + Environment.NewLine);

        }

        public void WritePerspectives()
        {
            txtProgress.AppendText( Progressstr + " Perspectives Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;
            ExcelWorkSheet = ExcelWorkBook.Sheets["Perspectives"];
            
            //OLAPCube.Perspectives[0].MeasureGroups[0].Measures[0].Measure.
            foreach(Perspective CubePerspective in OLAPCube.Perspectives)
            {
                foreach (PerspectiveDimension CubePerspectiveDim in CubePerspective.Dimensions)
                {
                    
                    foreach (PerspectiveAttribute CubePerspectiveAttribute in CubePerspectiveDim.Attributes)
                    {
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = CubePerspective.Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = OLAPCube.Dimensions.Find(CubePerspectiveDim.CubeDimensionID).Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = OLAPCube.Dimensions.Find(CubePerspectiveDim.CubeDimensionID).Attributes.Find(CubePerspectiveAttribute.AttributeID).Attribute.Name.ToString().Replace(CubePerspectiveDim.CubeDimensionID + ".", "");
                        ExcelSheetStartrow++;
                    }
                }

                
                /*
                foreach (MeasureGroup CubePerspectiveMeasureGroup in CubePerspective.MeasureGroups)
                {
                    foreach(Measure CubePerspectiveMeasure in CubePerspectiveMeasureGroup.Measures )
                    {
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = CubePerspective.Name;
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = "Measure";
                        ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = CubePerspectiveMeasureGroup.Name + "-" + CubePerspectiveMeasure.Name;
                        ExcelSheetStartrow++;

                    }
                }
                 * */
                
            }
            ExcelWorkBook.Save();

            txtProgress.AppendText( Progressstr + " Perspectives Completed " + Environment.NewLine);
        }

        public void WriteRole()
        {
            txtProgress.AppendText( Progressstr + " Roles Started " + Environment.NewLine);
            ExcelSheetStartrow = 2;
            ExcelWorkSheet = ExcelWorkBook.Sheets["Roles"];

            foreach(DatabasePermission dbPermission in OLAPDatabase.DatabasePermissions)
            {

                foreach(Dimension Dim in OLAPDatabase.Dimensions)
                {
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 1] = dbPermission.Role.Name;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 2] = dbPermission.Role.Description;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 3] = dbPermission.Administer;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 4] = dbPermission.Process;
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 5] = dbPermission.Read.ToString();
                    ExcelWorkSheet.Cells[ExcelSheetStartrow, 6] = Dim.Name;

                    if (Dim.DimensionPermissions.Count > 0)
                    {
                        if (Dim.DimensionPermissions[0].RoleID == dbPermission.RoleID)
                        {
                            ExcelWorkSheet.Cells[ExcelSheetStartrow, 7] = Dim.DimensionPermissions[0].AllowedRowsExpression;
                        }
                    }
                    ExcelSheetStartrow++;

                }
                
            }
            ExcelWorkBook.Save();

            txtProgress.AppendText( Progressstr + " Roles Completed " + Environment.NewLine);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
           txtOutputPath.Text =  System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\Output";
        }

        private void cmdSelectAll_Click(object sender, EventArgs e)
        {
            for(int i=0;i<=clstDBCubeName.Items.Count-1;i++)
            {

                clstDBCubeName.SetItemChecked(i,true); 
            }
        }

        private void cmdUnselectAll_Click(object sender, EventArgs e)
        {

            for (int i = 0; i <= clstDBCubeName.Items.Count - 1; i++)
            {

                clstDBCubeName.SetItemChecked(i, false);
            }         

        }

    }
}
