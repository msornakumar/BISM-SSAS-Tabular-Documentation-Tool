namespace SSASDocumentationTool_DescriptionEditor
{
    partial class frmSSASDocumentationTool
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSSASDocumentationTool));
            this.lblServerName = new System.Windows.Forms.Label();
            this.txtServerName = new System.Windows.Forms.TextBox();
            this.cmdGenerateDocument = new System.Windows.Forms.Button();
            this.txtProgress = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBoxSSASInstance = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cboCubeName = new System.Windows.Forms.ComboBox();
            this.lblCubeName = new System.Windows.Forms.Label();
            this.cboDatabaseName = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupConnection = new System.Windows.Forms.GroupBox();
            this.cmdConnect = new System.Windows.Forms.Button();
            this.checkBoxCurrentCreds = new System.Windows.Forms.CheckBox();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.lblUserName = new System.Windows.Forms.Label();
            this.textBoxUserName = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.progressGeneration = new System.Windows.Forms.ProgressBar();
            this.groupOutputConfig = new System.Windows.Forms.GroupBox();
            this.checkBoxOpenXL = new System.Windows.Forms.CheckBox();
            this.txtOutputPath = new System.Windows.Forms.TextBox();
            this.lblOutputPath = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupSSASServerType = new System.Windows.Forms.GroupBox();
            this.rdoMultiDimensional = new System.Windows.Forms.RadioButton();
            this.rdoTabular = new System.Windows.Forms.RadioButton();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupSSASInstanceType = new System.Windows.Forms.GroupBox();
            this.radioPBIDevInstance = new System.Windows.Forms.RadioButton();
            this.radioSSASTabularInstance = new System.Windows.Forms.RadioButton();
            this.groupPBIDevEnvInstance = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cboLocalHost = new System.Windows.Forms.ComboBox();
            this.lblLocalHost = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBoxSSASInstance.SuspendLayout();
            this.groupConnection.SuspendLayout();
            this.groupOutputConfig.SuspendLayout();
            this.groupSSASServerType.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupSSASInstanceType.SuspendLayout();
            this.groupPBIDevEnvInstance.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblServerName
            // 
            this.lblServerName.AutoSize = true;
            this.lblServerName.Location = new System.Drawing.Point(19, 16);
            this.lblServerName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblServerName.Name = "lblServerName";
            this.lblServerName.Size = new System.Drawing.Size(138, 17);
            this.lblServerName.TabIndex = 0;
            this.lblServerName.Text = "SSAS Server Type  :";
            // 
            // txtServerName
            // 
            this.txtServerName.Location = new System.Drawing.Point(147, 32);
            this.txtServerName.Margin = new System.Windows.Forms.Padding(4);
            this.txtServerName.Name = "txtServerName";
            this.txtServerName.Size = new System.Drawing.Size(390, 22);
            this.txtServerName.TabIndex = 1;
            // 
            // cmdGenerateDocument
            // 
            this.cmdGenerateDocument.FlatAppearance.BorderSize = 5;
            this.cmdGenerateDocument.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.cmdGenerateDocument.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdGenerateDocument.ForeColor = System.Drawing.Color.Black;
            this.cmdGenerateDocument.Location = new System.Drawing.Point(226, 432);
            this.cmdGenerateDocument.Margin = new System.Windows.Forms.Padding(4);
            this.cmdGenerateDocument.Name = "cmdGenerateDocument";
            this.cmdGenerateDocument.Size = new System.Drawing.Size(441, 36);
            this.cmdGenerateDocument.TabIndex = 7;
            this.cmdGenerateDocument.Text = "Generate";
            this.cmdGenerateDocument.UseVisualStyleBackColor = true;
            this.cmdGenerateDocument.Click += new System.EventHandler(this.cmdGenerateDocument_Click);
            // 
            // txtProgress
            // 
            this.txtProgress.Location = new System.Drawing.Point(11, 529);
            this.txtProgress.Margin = new System.Windows.Forms.Padding(4);
            this.txtProgress.Multiline = true;
            this.txtProgress.Name = "txtProgress";
            this.txtProgress.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtProgress.Size = new System.Drawing.Size(987, 197);
            this.txtProgress.TabIndex = 5;
            this.txtProgress.TextChanged += new System.EventHandler(this.txtProgress_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(11, 466);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 17);
            this.label2.TabIndex = 9;
            this.label2.Text = "Progress :";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // groupBoxSSASInstance
            // 
            this.groupBoxSSASInstance.Controls.Add(this.label5);
            this.groupBoxSSASInstance.Controls.Add(this.cboCubeName);
            this.groupBoxSSASInstance.Controls.Add(this.txtServerName);
            this.groupBoxSSASInstance.Controls.Add(this.lblCubeName);
            this.groupBoxSSASInstance.Controls.Add(this.cboDatabaseName);
            this.groupBoxSSASInstance.Controls.Add(this.label3);
            this.groupBoxSSASInstance.Controls.Add(this.groupConnection);
            this.groupBoxSSASInstance.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBoxSSASInstance.Location = new System.Drawing.Point(11, 142);
            this.groupBoxSSASInstance.Margin = new System.Windows.Forms.Padding(13, 12, 13, 12);
            this.groupBoxSSASInstance.Name = "groupBoxSSASInstance";
            this.groupBoxSSASInstance.Padding = new System.Windows.Forms.Padding(4);
            this.groupBoxSSASInstance.Size = new System.Drawing.Size(988, 161);
            this.groupBoxSSASInstance.TabIndex = 12;
            this.groupBoxSSASInstance.TabStop = false;
            this.groupBoxSSASInstance.Text = "SSAS Tabular Instance";
            this.groupBoxSSASInstance.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(6, 36);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(103, 17);
            this.label5.TabIndex = 17;
            this.label5.Text = "Server Name  :";
            // 
            // cboCubeName
            // 
            this.cboCubeName.FormattingEnabled = true;
            this.cboCubeName.Location = new System.Drawing.Point(147, 110);
            this.cboCubeName.Margin = new System.Windows.Forms.Padding(4);
            this.cboCubeName.Name = "cboCubeName";
            this.cboCubeName.Size = new System.Drawing.Size(390, 24);
            this.cboCubeName.TabIndex = 4;
            this.cboCubeName.SelectedIndexChanged += new System.EventHandler(this.cboCubeName_SelectedIndexChanged);
            // 
            // lblCubeName
            // 
            this.lblCubeName.AutoSize = true;
            this.lblCubeName.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCubeName.Location = new System.Drawing.Point(6, 112);
            this.lblCubeName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblCubeName.Name = "lblCubeName";
            this.lblCubeName.Size = new System.Drawing.Size(94, 17);
            this.lblCubeName.TabIndex = 15;
            this.lblCubeName.Text = "Cube Name  :";
            // 
            // cboDatabaseName
            // 
            this.cboDatabaseName.FormattingEnabled = true;
            this.cboDatabaseName.Location = new System.Drawing.Point(147, 70);
            this.cboDatabaseName.Margin = new System.Windows.Forms.Padding(4);
            this.cboDatabaseName.Name = "cboDatabaseName";
            this.cboDatabaseName.Size = new System.Drawing.Size(390, 24);
            this.cboDatabaseName.Sorted = true;
            this.cboDatabaseName.TabIndex = 3;
            this.cboDatabaseName.SelectedIndexChanged += new System.EventHandler(this.cboDatabaseName_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(6, 74);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(122, 17);
            this.label3.TabIndex = 13;
            this.label3.Text = "Database Name  :";
            // 
            // groupConnection
            // 
            this.groupConnection.Controls.Add(this.cmdConnect);
            this.groupConnection.Controls.Add(this.checkBoxCurrentCreds);
            this.groupConnection.Controls.Add(this.textBoxPassword);
            this.groupConnection.Controls.Add(this.lblUserName);
            this.groupConnection.Controls.Add(this.textBoxUserName);
            this.groupConnection.Controls.Add(this.label7);
            this.groupConnection.Location = new System.Drawing.Point(560, 9);
            this.groupConnection.Name = "groupConnection";
            this.groupConnection.Size = new System.Drawing.Size(421, 143);
            this.groupConnection.TabIndex = 25;
            this.groupConnection.TabStop = false;
            this.groupConnection.Text = "Connection Credentials";
            // 
            // cmdConnect
            // 
            this.cmdConnect.Location = new System.Drawing.Point(291, 112);
            this.cmdConnect.Margin = new System.Windows.Forms.Padding(4);
            this.cmdConnect.Name = "cmdConnect";
            this.cmdConnect.Size = new System.Drawing.Size(117, 25);
            this.cmdConnect.TabIndex = 24;
            this.cmdConnect.Text = "Connect";
            this.cmdConnect.UseVisualStyleBackColor = true;
            this.cmdConnect.Click += new System.EventHandler(this.cmdConnect_Click);
            // 
            // checkBoxCurrentCreds
            // 
            this.checkBoxCurrentCreds.AutoSize = true;
            this.checkBoxCurrentCreds.Checked = true;
            this.checkBoxCurrentCreds.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxCurrentCreds.Location = new System.Drawing.Point(24, 21);
            this.checkBoxCurrentCreds.Name = "checkBoxCurrentCreds";
            this.checkBoxCurrentCreds.Size = new System.Drawing.Size(241, 21);
            this.checkBoxCurrentCreds.TabIndex = 23;
            this.checkBoxCurrentCreds.Text = "Use Current Windows Credentials";
            this.checkBoxCurrentCreds.UseVisualStyleBackColor = true;
            this.checkBoxCurrentCreds.CheckedChanged += new System.EventHandler(this.checkBoxCurrentCreds_CheckedChanged);
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Enabled = false;
            this.textBoxPassword.Location = new System.Drawing.Point(165, 84);
            this.textBoxPassword.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.PasswordChar = '*';
            this.textBoxPassword.Size = new System.Drawing.Size(246, 22);
            this.textBoxPassword.TabIndex = 22;
            // 
            // lblUserName
            // 
            this.lblUserName.AutoSize = true;
            this.lblUserName.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUserName.Location = new System.Drawing.Point(24, 49);
            this.lblUserName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(91, 17);
            this.lblUserName.TabIndex = 21;
            this.lblUserName.Text = "User Name  :";
            // 
            // textBoxUserName
            // 
            this.textBoxUserName.Enabled = false;
            this.textBoxUserName.Location = new System.Drawing.Point(165, 46);
            this.textBoxUserName.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxUserName.Name = "textBoxUserName";
            this.textBoxUserName.Size = new System.Drawing.Size(246, 22);
            this.textBoxUserName.TabIndex = 18;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(24, 87);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(89, 17);
            this.label7.TabIndex = 20;
            this.label7.Text = "Password    :";
            // 
            // progressGeneration
            // 
            this.progressGeneration.Location = new System.Drawing.Point(11, 495);
            this.progressGeneration.Margin = new System.Windows.Forms.Padding(4);
            this.progressGeneration.Maximum = 10;
            this.progressGeneration.Name = "progressGeneration";
            this.progressGeneration.Size = new System.Drawing.Size(987, 28);
            this.progressGeneration.TabIndex = 15;
            // 
            // groupOutputConfig
            // 
            this.groupOutputConfig.Controls.Add(this.checkBoxOpenXL);
            this.groupOutputConfig.Controls.Add(this.txtOutputPath);
            this.groupOutputConfig.Controls.Add(this.lblOutputPath);
            this.groupOutputConfig.Controls.Add(this.txtFileName);
            this.groupOutputConfig.Controls.Add(this.label6);
            this.groupOutputConfig.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupOutputConfig.Location = new System.Drawing.Point(11, 302);
            this.groupOutputConfig.Name = "groupOutputConfig";
            this.groupOutputConfig.Size = new System.Drawing.Size(987, 123);
            this.groupOutputConfig.TabIndex = 16;
            this.groupOutputConfig.TabStop = false;
            this.groupOutputConfig.Text = "Output Configuration";
            // 
            // checkBoxOpenXL
            // 
            this.checkBoxOpenXL.AutoSize = true;
            this.checkBoxOpenXL.Location = new System.Drawing.Point(147, 94);
            this.checkBoxOpenXL.Name = "checkBoxOpenXL";
            this.checkBoxOpenXL.Size = new System.Drawing.Size(207, 21);
            this.checkBoxOpenXL.TabIndex = 25;
            this.checkBoxOpenXL.Text = "Open Excel after completion";
            this.checkBoxOpenXL.UseVisualStyleBackColor = true;
            // 
            // txtOutputPath
            // 
            this.txtOutputPath.Location = new System.Drawing.Point(147, 23);
            this.txtOutputPath.Margin = new System.Windows.Forms.Padding(4);
            this.txtOutputPath.Name = "txtOutputPath";
            this.txtOutputPath.Size = new System.Drawing.Size(644, 22);
            this.txtOutputPath.TabIndex = 21;
            // 
            // lblOutputPath
            // 
            this.lblOutputPath.AutoSize = true;
            this.lblOutputPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOutputPath.Location = new System.Drawing.Point(6, 23);
            this.lblOutputPath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblOutputPath.Name = "lblOutputPath";
            this.lblOutputPath.Size = new System.Drawing.Size(96, 17);
            this.lblOutputPath.TabIndex = 23;
            this.lblOutputPath.Text = "Output Path  :";
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(147, 61);
            this.txtFileName.Margin = new System.Windows.Forms.Padding(4);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(644, 22);
            this.txtFileName.TabIndex = 22;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(6, 61);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(83, 17);
            this.label6.TabIndex = 24;
            this.label6.Text = "File Name  :";
            // 
            // groupSSASServerType
            // 
            this.groupSSASServerType.Controls.Add(this.rdoMultiDimensional);
            this.groupSSASServerType.Controls.Add(this.rdoTabular);
            this.groupSSASServerType.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupSSASServerType.ForeColor = System.Drawing.Color.Black;
            this.groupSSASServerType.Location = new System.Drawing.Point(11, 3);
            this.groupSSASServerType.Margin = new System.Windows.Forms.Padding(4);
            this.groupSSASServerType.Name = "groupSSASServerType";
            this.groupSSASServerType.Padding = new System.Windows.Forms.Padding(4);
            this.groupSSASServerType.Size = new System.Drawing.Size(822, 63);
            this.groupSSASServerType.TabIndex = 19;
            this.groupSSASServerType.TabStop = false;
            this.groupSSASServerType.Text = "SSAS Server Type";
            // 
            // rdoMultiDimensional
            // 
            this.rdoMultiDimensional.AutoSize = true;
            this.rdoMultiDimensional.Enabled = false;
            this.rdoMultiDimensional.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoMultiDimensional.Location = new System.Drawing.Point(382, 23);
            this.rdoMultiDimensional.Margin = new System.Windows.Forms.Padding(4);
            this.rdoMultiDimensional.Name = "rdoMultiDimensional";
            this.rdoMultiDimensional.Size = new System.Drawing.Size(135, 21);
            this.rdoMultiDimensional.TabIndex = 1;
            this.rdoMultiDimensional.Text = "MultiDimensional";
            this.rdoMultiDimensional.UseVisualStyleBackColor = true;
            // 
            // rdoTabular
            // 
            this.rdoTabular.AutoSize = true;
            this.rdoTabular.Checked = true;
            this.rdoTabular.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoTabular.Location = new System.Drawing.Point(147, 22);
            this.rdoTabular.Margin = new System.Windows.Forms.Padding(4);
            this.rdoTabular.Name = "rdoTabular";
            this.rdoTabular.Size = new System.Drawing.Size(78, 21);
            this.rdoTabular.TabIndex = 0;
            this.rdoTabular.TabStop = true;
            this.rdoTabular.Text = "Tabular";
            this.rdoTabular.UseVisualStyleBackColor = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(864, 17);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(113, 119);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 20;
            this.pictureBox1.TabStop = false;
            // 
            // groupSSASInstanceType
            // 
            this.groupSSASInstanceType.Controls.Add(this.radioPBIDevInstance);
            this.groupSSASInstanceType.Controls.Add(this.radioSSASTabularInstance);
            this.groupSSASInstanceType.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupSSASInstanceType.Location = new System.Drawing.Point(11, 74);
            this.groupSSASInstanceType.Margin = new System.Windows.Forms.Padding(10);
            this.groupSSASInstanceType.Name = "groupSSASInstanceType";
            this.groupSSASInstanceType.Padding = new System.Windows.Forms.Padding(4);
            this.groupSSASInstanceType.Size = new System.Drawing.Size(822, 62);
            this.groupSSASInstanceType.TabIndex = 21;
            this.groupSSASInstanceType.TabStop = false;
            this.groupSSASInstanceType.Text = "SSAS Instance Type";
            // 
            // radioPBIDevInstance
            // 
            this.radioPBIDevInstance.AutoSize = true;
            this.radioPBIDevInstance.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioPBIDevInstance.Location = new System.Drawing.Point(381, 25);
            this.radioPBIDevInstance.Margin = new System.Windows.Forms.Padding(4);
            this.radioPBIDevInstance.Name = "radioPBIDevInstance";
            this.radioPBIDevInstance.Size = new System.Drawing.Size(206, 21);
            this.radioPBIDevInstance.TabIndex = 1;
            this.radioPBIDevInstance.Text = "Power BI \\ Dev Env Instance";
            this.radioPBIDevInstance.UseVisualStyleBackColor = true;
            this.radioPBIDevInstance.CheckedChanged += new System.EventHandler(this.radioPBIDevInstance_CheckedChanged);
            // 
            // radioSSASTabularInstance
            // 
            this.radioSSASTabularInstance.AutoSize = true;
            this.radioSSASTabularInstance.Checked = true;
            this.radioSSASTabularInstance.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioSSASTabularInstance.Location = new System.Drawing.Point(147, 24);
            this.radioSSASTabularInstance.Margin = new System.Windows.Forms.Padding(4);
            this.radioSSASTabularInstance.Name = "radioSSASTabularInstance";
            this.radioSSASTabularInstance.Size = new System.Drawing.Size(175, 21);
            this.radioSSASTabularInstance.TabIndex = 0;
            this.radioSSASTabularInstance.TabStop = true;
            this.radioSSASTabularInstance.Text = "SSAS Tabular Instance";
            this.radioSSASTabularInstance.UseVisualStyleBackColor = true;
            this.radioSSASTabularInstance.CheckedChanged += new System.EventHandler(this.radioSSASTabularInstance_CheckedChanged);
            // 
            // groupPBIDevEnvInstance
            // 
            this.groupPBIDevEnvInstance.Controls.Add(this.label4);
            this.groupPBIDevEnvInstance.Controls.Add(this.cboLocalHost);
            this.groupPBIDevEnvInstance.Controls.Add(this.lblLocalHost);
            this.groupPBIDevEnvInstance.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupPBIDevEnvInstance.Location = new System.Drawing.Point(11, 142);
            this.groupPBIDevEnvInstance.Name = "groupPBIDevEnvInstance";
            this.groupPBIDevEnvInstance.Size = new System.Drawing.Size(988, 153);
            this.groupPBIDevEnvInstance.TabIndex = 22;
            this.groupPBIDevEnvInstance.TabStop = false;
            this.groupPBIDevEnvInstance.Text = "PowerBI \\ Dev Env Instances";
            this.groupPBIDevEnvInstance.Visible = false;
            this.groupPBIDevEnvInstance.Enter += new System.EventHandler(this.groupBox5_Enter);
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(144, 99);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(564, 47);
            this.label4.TabIndex = 25;
            this.label4.Text = "Lists the Local Host models for open Power BI files && SSAS Visual Studio Dev Env" +
    "ironment open and using Integrated mode as Work Space";
            // 
            // cboLocalHost
            // 
            this.cboLocalHost.FormattingEnabled = true;
            this.cboLocalHost.Location = new System.Drawing.Point(147, 67);
            this.cboLocalHost.Margin = new System.Windows.Forms.Padding(4);
            this.cboLocalHost.Name = "cboLocalHost";
            this.cboLocalHost.Size = new System.Drawing.Size(644, 24);
            this.cboLocalHost.Sorted = true;
            this.cboLocalHost.TabIndex = 23;
            this.cboLocalHost.DropDown += new System.EventHandler(this.cboLocalHost_DropDown);
            this.cboLocalHost.SelectedIndexChanged += new System.EventHandler(this.cboLocalHost_SelectedIndexChanged);
            // 
            // lblLocalHost
            // 
            this.lblLocalHost.AutoSize = true;
            this.lblLocalHost.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLocalHost.Location = new System.Drawing.Point(10, 67);
            this.lblLocalHost.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblLocalHost.Name = "lblLocalHost";
            this.lblLocalHost.Size = new System.Drawing.Size(87, 17);
            this.lblLocalHost.TabIndex = 24;
            this.lblLocalHost.Text = "Local Host  :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 75);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 17);
            this.label1.TabIndex = 23;
            this.label1.Text = "SSAS Instance Type  :";
            // 
            // frmSSASDocumentationTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1002, 734);
            this.Controls.Add(this.groupSSASInstanceType);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.groupSSASServerType);
            this.Controls.Add(this.groupOutputConfig);
            this.Controls.Add(this.progressGeneration);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtProgress);
            this.Controls.Add(this.cmdGenerateDocument);
            this.Controls.Add(this.lblServerName);
            this.Controls.Add(this.groupBoxSSASInstance);
            this.Controls.Add(this.groupPBIDevEnvInstance);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "frmSSASDocumentationTool";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Tag = " ";
            this.Text = "BISM SSAS Documentation Tool";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.groupBoxSSASInstance.ResumeLayout(false);
            this.groupBoxSSASInstance.PerformLayout();
            this.groupConnection.ResumeLayout(false);
            this.groupConnection.PerformLayout();
            this.groupOutputConfig.ResumeLayout(false);
            this.groupOutputConfig.PerformLayout();
            this.groupSSASServerType.ResumeLayout(false);
            this.groupSSASServerType.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupSSASInstanceType.ResumeLayout(false);
            this.groupSSASInstanceType.PerformLayout();
            this.groupPBIDevEnvInstance.ResumeLayout(false);
            this.groupPBIDevEnvInstance.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblServerName;
        private System.Windows.Forms.TextBox txtServerName;
        private System.Windows.Forms.Button cmdGenerateDocument;
        private System.Windows.Forms.TextBox txtProgress;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBoxSSASInstance;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboDatabaseName;
        private System.Windows.Forms.ComboBox cboCubeName;
        private System.Windows.Forms.Label lblCubeName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ProgressBar progressGeneration;
        private System.Windows.Forms.GroupBox groupOutputConfig;
        private System.Windows.Forms.TextBox txtOutputPath;
        private System.Windows.Forms.Label lblOutputPath;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupSSASServerType;
        private System.Windows.Forms.RadioButton rdoMultiDimensional;
        private System.Windows.Forms.RadioButton rdoTabular;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.GroupBox groupSSASInstanceType;
        private System.Windows.Forms.RadioButton radioPBIDevInstance;
        private System.Windows.Forms.RadioButton radioSSASTabularInstance;
        private System.Windows.Forms.GroupBox groupPBIDevEnvInstance;
        private System.Windows.Forms.ComboBox cboLocalHost;
        private System.Windows.Forms.Label lblLocalHost;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupConnection;
        private System.Windows.Forms.Button cmdConnect;
        private System.Windows.Forms.CheckBox checkBoxCurrentCreds;
        private System.Windows.Forms.TextBox textBoxPassword;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.TextBox textBoxUserName;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox checkBoxOpenXL;
        private System.Windows.Forms.Label label4;
    }
}