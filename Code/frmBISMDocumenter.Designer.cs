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
            this.lblServerName = new System.Windows.Forms.Label();
            this.txtServerName = new System.Windows.Forms.TextBox();
            this.cmdConnect = new System.Windows.Forms.Button();
            this.cmdGenerateDocument = new System.Windows.Forms.Button();
            this.txtProgress = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtOutputPath = new System.Windows.Forms.TextBox();
            this.lblOutputPath = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rdoMultiDimensional = new System.Windows.Forms.RadioButton();
            this.rdoTabular = new System.Windows.Forms.RadioButton();
            this.label5 = new System.Windows.Forms.Label();
            this.cboCubeName = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cboDatabaseName = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.progressGeneration = new System.Windows.Forms.ProgressBar();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblServerName
            // 
            this.lblServerName.AutoSize = true;
            this.lblServerName.Location = new System.Drawing.Point(8, 15);
            this.lblServerName.Name = "lblServerName";
            this.lblServerName.Size = new System.Drawing.Size(74, 13);
            this.lblServerName.TabIndex = 0;
            this.lblServerName.Text = "Server Type  :";
            // 
            // txtServerName
            // 
            this.txtServerName.Location = new System.Drawing.Point(112, 58);
            this.txtServerName.Name = "txtServerName";
            this.txtServerName.Size = new System.Drawing.Size(484, 20);
            this.txtServerName.TabIndex = 1;
            // 
            // cmdConnect
            // 
            this.cmdConnect.Location = new System.Drawing.Point(602, 58);
            this.cmdConnect.Name = "cmdConnect";
            this.cmdConnect.Size = new System.Drawing.Size(129, 20);
            this.cmdConnect.TabIndex = 2;
            this.cmdConnect.Text = "Connect";
            this.cmdConnect.UseVisualStyleBackColor = true;
            this.cmdConnect.Click += new System.EventHandler(this.cmdConnect_Click);
            // 
            // cmdGenerateDocument
            // 
            this.cmdGenerateDocument.FlatAppearance.BorderSize = 5;
            this.cmdGenerateDocument.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.cmdGenerateDocument.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdGenerateDocument.ForeColor = System.Drawing.Color.Black;
            this.cmdGenerateDocument.Location = new System.Drawing.Point(178, 213);
            this.cmdGenerateDocument.Name = "cmdGenerateDocument";
            this.cmdGenerateDocument.Size = new System.Drawing.Size(331, 29);
            this.cmdGenerateDocument.TabIndex = 7;
            this.cmdGenerateDocument.Text = "Generate";
            this.cmdGenerateDocument.UseVisualStyleBackColor = true;
            this.cmdGenerateDocument.Click += new System.EventHandler(this.cmdGenerateDocument_Click);
            // 
            // txtProgress
            // 
            this.txtProgress.Location = new System.Drawing.Point(1, 303);
            this.txtProgress.Multiline = true;
            this.txtProgress.Name = "txtProgress";
            this.txtProgress.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtProgress.Size = new System.Drawing.Size(741, 181);
            this.txtProgress.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 255);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Progress :";
            // 
            // txtOutputPath
            // 
            this.txtOutputPath.Location = new System.Drawing.Point(112, 146);
            this.txtOutputPath.Name = "txtOutputPath";
            this.txtOutputPath.Size = new System.Drawing.Size(484, 20);
            this.txtOutputPath.TabIndex = 5;
            // 
            // lblOutputPath
            // 
            this.lblOutputPath.AutoSize = true;
            this.lblOutputPath.Location = new System.Drawing.Point(8, 149);
            this.lblOutputPath.Name = "lblOutputPath";
            this.lblOutputPath.Size = new System.Drawing.Size(73, 13);
            this.lblOutputPath.TabIndex = 10;
            this.lblOutputPath.Text = "Output Path  :";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtFileName);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.cboCubeName);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.cboDatabaseName);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Location = new System.Drawing.Point(0, 2);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(10);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(741, 205);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(112, 170);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(484, 20);
            this.txtFileName.TabIndex = 6;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(8, 173);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(63, 13);
            this.label6.TabIndex = 20;
            this.label6.Text = "File Name  :";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rdoMultiDimensional);
            this.groupBox2.Controls.Add(this.rdoTabular);
            this.groupBox2.Location = new System.Drawing.Point(97, 8);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(633, 42);
            this.groupBox2.TabIndex = 18;
            this.groupBox2.TabStop = false;
            // 
            // rdoMultiDimensional
            // 
            this.rdoMultiDimensional.AutoSize = true;
            this.rdoMultiDimensional.Enabled = false;
            this.rdoMultiDimensional.Location = new System.Drawing.Point(267, 12);
            this.rdoMultiDimensional.Name = "rdoMultiDimensional";
            this.rdoMultiDimensional.Size = new System.Drawing.Size(104, 17);
            this.rdoMultiDimensional.TabIndex = 1;
            this.rdoMultiDimensional.Text = "MultiDimensional";
            this.rdoMultiDimensional.UseVisualStyleBackColor = true;
            // 
            // rdoTabular
            // 
            this.rdoTabular.AutoSize = true;
            this.rdoTabular.Checked = true;
            this.rdoTabular.Location = new System.Drawing.Point(91, 11);
            this.rdoTabular.Name = "rdoTabular";
            this.rdoTabular.Size = new System.Drawing.Size(61, 17);
            this.rdoTabular.TabIndex = 0;
            this.rdoTabular.TabStop = true;
            this.rdoTabular.Text = "Tabular";
            this.rdoTabular.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 59);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(78, 13);
            this.label5.TabIndex = 17;
            this.label5.Text = "Server Name  :";
            // 
            // cboCubeName
            // 
            this.cboCubeName.FormattingEnabled = true;
            this.cboCubeName.Location = new System.Drawing.Point(113, 115);
            this.cboCubeName.Name = "cboCubeName";
            this.cboCubeName.Size = new System.Drawing.Size(484, 21);
            this.cboCubeName.TabIndex = 4;
            this.cboCubeName.SelectedIndexChanged += new System.EventHandler(this.cboCubeName_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Cube Name  :";
            // 
            // cboDatabaseName
            // 
            this.cboDatabaseName.FormattingEnabled = true;
            this.cboDatabaseName.Location = new System.Drawing.Point(113, 87);
            this.cboDatabaseName.Name = "cboDatabaseName";
            this.cboDatabaseName.Size = new System.Drawing.Size(484, 21);
            this.cboDatabaseName.Sorted = true;
            this.cboDatabaseName.TabIndex = 3;
            this.cboDatabaseName.SelectedIndexChanged += new System.EventHandler(this.cboDatabaseName_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 92);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Database Name  :";
            // 
            // progressGeneration
            // 
            this.progressGeneration.Location = new System.Drawing.Point(1, 274);
            this.progressGeneration.Maximum = 10;
            this.progressGeneration.Name = "progressGeneration";
            this.progressGeneration.Size = new System.Drawing.Size(740, 23);
            this.progressGeneration.TabIndex = 15;
            // 
            // frmSSASDocumentationTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(741, 494);
            this.Controls.Add(this.progressGeneration);
            this.Controls.Add(this.txtOutputPath);
            this.Controls.Add(this.lblOutputPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtProgress);
            this.Controls.Add(this.cmdGenerateDocument);
            this.Controls.Add(this.cmdConnect);
            this.Controls.Add(this.txtServerName);
            this.Controls.Add(this.lblServerName);
            this.Controls.Add(this.groupBox1);
            this.Name = "frmSSASDocumentationTool";
            this.Tag = " ";
            this.Text = "SSAS Documentation Tool";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblServerName;
        private System.Windows.Forms.TextBox txtServerName;
        private System.Windows.Forms.Button cmdConnect;
        private System.Windows.Forms.Button cmdGenerateDocument;
        private System.Windows.Forms.TextBox txtProgress;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtOutputPath;
        private System.Windows.Forms.Label lblOutputPath;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboDatabaseName;
        private System.Windows.Forms.ComboBox cboCubeName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rdoMultiDimensional;
        private System.Windows.Forms.RadioButton rdoTabular;
        private System.Windows.Forms.ProgressBar progressGeneration;
    }
}