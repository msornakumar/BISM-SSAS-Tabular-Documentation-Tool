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
using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
//using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;


namespace SSASDocumentationTool_DescriptionEditor
{
    public partial class frmSSASDocumentationTool : Form
    {

        string AppPath;
        string OutputPath;

      
        String OLAPServerName;
        String OLAPDBName;
        String OLAPCubeName;

        AMO.Server OLAPServer;
        AMO.Database OLAPDatabase;
        AMO.Cube OLAPCube;

        TOM.Server TOMServer;
        TOM.Database TOMDb;

        string Progressstr;

        int TabularCompatibilityLevel;

        
        

        public frmSSASDocumentationTool()
        {
            InitializeComponent();
        }

        private void cmdConnect_Click(object sender, EventArgs e)
        {

            try
            {
                cboDatabaseName.Items.Clear();
                cboDatabaseName.Text = "";
                cboCubeName.Items.Clear();
                cboCubeName.Text = "";

                String ConnStr;
                OLAPServerName = txtServerName.Text;
                txtProgress.AppendText("");

                if (checkBoxCurrentCreds.Checked == true)
                {
                    ConnStr = "Provider=MSOLAP;Data Source=" + OLAPServerName + ";";
                }
                else
                {
                    // Provider = MSOLAP.8; Persist Security Info = True; User ID = sg3\msorn; Initial Catalog = SSASTOM; Data Source = sg3\sql2017; MDX Compatibility = 1; Safety Options = 2; MDX Missing Member Mode = Error; Update Isolation Level = 2
                    ConnStr = "Provider=MSOLAP;Data Source=" + OLAPServerName + ";User Id = " + textBoxUserName.Text + "; Password = " + textBoxPassword.Text + ";";
                }
                //Initial Catalog=Adventure Works DW 2008R2;"; 
                OLAPServer = new AMO.Server();

                TOMServer = new TOM.Server();

                
                

                OLAPServer.Connect(ConnStr);
                TOMServer.Connect(ConnStr);
                
                Console.WriteLine("ServerName : " + OLAPServerName);

                cboDatabaseName.Items.Clear();
                foreach (AMO.Database OLAPDatabase in OLAPServer.Databases)
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

            if (radioSSASTabularInstance.Checked ==true)
            {
                if (txtServerName.Text.Trim() == "")
                {
                    MessageBox.Show("Please enter ServerName and Click connect with proper credential selection", "BISMDocumenter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (cboDatabaseName.Text.Trim() == "")
                {
                    MessageBox.Show("Please select a Database Name", "BISMDocumenter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (cboCubeName.Enabled == true &&  cboCubeName.Text.Trim() == "")
                {
                    MessageBox.Show("Please select a Cube Name", "BISMDocumenter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            if (radioPBIDevInstance.Checked == true)
            {
                if (cboLocalHost.Text.Trim() == "")
                {
                    MessageBox.Show("Please select a LocalHost Name", "BISMDocumenter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

            }

            if (txtFileName.Text.Trim() == "")
            {
                MessageBox.Show("Please enter a File Name", "BISMDocumenter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                progressGeneration.Maximum = 30;
                progressGeneration.Value = 0;
                String Filename;

                //frmSSASDocumentationTool.ActiveForm.Height = 430;

                AppPath = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                
                if (!System.IO.Directory.Exists(txtOutputPath.Text))
                {
                    System.IO.Directory.CreateDirectory(txtOutputPath.Text);
                }

                OutputPath = txtOutputPath.Text;
                

                


                if (radioPBIDevInstance.Checked== true)
                {

                    OLAPServerName = cboLocalHost.SelectedItem.ToString();
                    OLAPServerName = OLAPServerName.Substring(0, OLAPServerName.IndexOf(".") );
                    OLAPServerName = OLAPServerName.Replace("Devenv -", "").Trim();
                    OLAPServerName = OLAPServerName.Replace("PowerBI -", "").Trim();
                    OLAPDBName = "";
                    TabularCompatibilityLevel = 1200;
                    txtFileName.Text = cboLocalHost.SelectedItem.ToString().Replace(OLAPServerName + ".","") ;
                }

                if (radioSSASTabularInstance.Checked == true)
                {
                    OLAPServerName = txtServerName.Text;
                    OLAPDBName = cboDatabaseName.SelectedItem.ToString();

                }

                Filename = txtFileName.Text + ".xlsx";


                if (TabularCompatibilityLevel < 1200)
                {
                    OLAPCubeName = cboCubeName.SelectedItem.ToString(); // DBCubeName.Substring(DBCubeName.IndexOf("-") + 1, DBCubeName.Length - (DBCubeName.IndexOf("-")+1));
                    BISMDocumenterAMO.BISMDocumenterCls BISMAMODoc = new BISMDocumenterAMO.BISMDocumenterCls();
                    BISMAMODoc.GenerateDocument(OLAPServerName, OLAPDBName, OLAPCubeName, OutputPath,Filename,txtProgress,checkBoxOpenXL.Checked);
                }
                else
                {
                    OLAPCubeName = "";
                    BISMDocumenterTOM.BISMDocumenterCls BISMTOMDoc = new BISMDocumenterTOM.BISMDocumenterCls();
                    BISMTOMDoc.GenerateDocument(OLAPServerName, OLAPDBName, OLAPCubeName, OutputPath,Filename, txtProgress, checkBoxOpenXL.Checked);
                }


                
                //frmSSASDocumentationTool.ActiveForm.Height = 390;
                progressGeneration.Value = progressGeneration.Maximum;

                if (!checkBoxOpenXL.Checked)
                {
                    MessageBox.Show("Generation Completed Succesfully", "SSAS Documentation Tool", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            
            catch(Exception err)
            {

                string errormsg = err.ToString();
                txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                txtProgress.AppendText("Error Occured" + Environment.NewLine);
                txtProgress.AppendText("--------------------------------------------------------------------------------------" + Environment.NewLine);
                txtProgress.AppendText(err.ToString() + Environment.NewLine);
                MessageBox.Show(errormsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                /*
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
                */

            }
            finally
            {
            /*
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLApp);
                GC.Collect();
                */
            }
            

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
            
            TOMDb =TOMServer.Databases[(cboDatabaseName.SelectedItem as ComboboxItem).Value.ToString()];

            TabularCompatibilityLevel = TOMDb.CompatibilityLevel;
            if (TabularCompatibilityLevel < 1200)
            {
                cboCubeName.Enabled = true;
                foreach (AMO.Cube OLAPCube in OLAPDatabase.Cubes)
                {

                    cboCubeName.Items.Add(OLAPCube.Name);

                }

                if (cboCubeName.Items.Count == 1)
                {
                    cboCubeName.SelectedItem = cboCubeName.Items[0];
                }
            }
            else
            {
                cboCubeName.Enabled= false;
                txtFileName.Text = cboDatabaseName.SelectedItem.ToString() ;
            }

        }

        private void cboCubeName_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtFileName.Text = cboDatabaseName.SelectedItem.ToString() + "-" + cboCubeName.SelectedItem.ToString();
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

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void cboLocalHost_SelectedIndexChanged(object sender, EventArgs e)
        {

            txtFileName.Text = cboLocalHost.SelectedItem.ToString();


        }

        private void cboLocalHost_DropDown(object sender, EventArgs e)
        {
            
            cboLocalHost.Items.Clear();
            PowerBIHelper.PowerBIHelper.Refresh();

            foreach (PowerBIHelper.PowerBIInstance PIX in PowerBIHelper.PowerBIHelper.Instances)
            {
                cboLocalHost.Items.Add(PIX.Icon.ToString() + " - localhost:" + PIX.Port + "." + PIX.Name);
            }

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void radioSSASTabularInstance_CheckedChanged(object sender, EventArgs e)
        {
            groupBoxSSASInstance.Visible = true;
            groupPBIDevEnvInstance.Visible = false;
            txtFileName.Text = "";
        }

        private void radioPBIDevInstance_CheckedChanged(object sender, EventArgs e)
        {

            groupBoxSSASInstance.Visible = false;
            groupPBIDevEnvInstance.Visible = true;
            txtFileName.Text = "";
        }

        private void cboLocalHost_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void txtProgress_TextChanged(object sender, EventArgs e)
        {
            if (progressGeneration.Value < progressGeneration.Maximum)
            {
                progressGeneration.Value = progressGeneration.Value + 1;
            }
        }

        private void checkBoxCurrentCreds_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxCurrentCreds.CheckState==CheckState.Checked)
            {
                textBoxUserName.Enabled = false;
                textBoxPassword.Enabled = false;
            }
            else
            {
                textBoxUserName.Enabled = true;
                textBoxPassword.Enabled = true;
            }
        }
    }
}
