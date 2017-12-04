namespace BISMDocumentor_DescriptionEditor
{
    partial class frmBISMDocumentor
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
            this.clstDBCubeName = new System.Windows.Forms.CheckedListBox();
            this.cmdGenerateDocument = new System.Windows.Forms.Button();
            this.txtProgress = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cmdSelectAll = new System.Windows.Forms.Button();
            this.cmdUnselectAll = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtOutputPath = new System.Windows.Forms.TextBox();
            this.lblOutputPath = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.SuspendLayout();
            // 
            // lblServerName
            // 
            this.lblServerName.AutoSize = true;
            this.lblServerName.Location = new System.Drawing.Point(8, 15);
            this.lblServerName.Name = "lblServerName";
            this.lblServerName.Size = new System.Drawing.Size(75, 13);
            this.lblServerName.TabIndex = 0;
            this.lblServerName.Text = "ServerName  :";
            // 
            // txtServerName
            // 
            this.txtServerName.Location = new System.Drawing.Point(112, 12);
            this.txtServerName.Name = "txtServerName";
            this.txtServerName.Size = new System.Drawing.Size(484, 20);
            this.txtServerName.TabIndex = 1;
            // 
            // cmdConnect
            // 
            this.cmdConnect.Location = new System.Drawing.Point(602, 11);
            this.cmdConnect.Name = "cmdConnect";
            this.cmdConnect.Size = new System.Drawing.Size(129, 20);
            this.cmdConnect.TabIndex = 2;
            this.cmdConnect.Text = "Connect";
            this.cmdConnect.UseVisualStyleBackColor = true;
            this.cmdConnect.Click += new System.EventHandler(this.cmdConnect_Click);
            // 
            // clstDBCubeName
            // 
            this.clstDBCubeName.FormattingEnabled = true;
            this.clstDBCubeName.Location = new System.Drawing.Point(112, 48);
            this.clstDBCubeName.Name = "clstDBCubeName";
            this.clstDBCubeName.Size = new System.Drawing.Size(484, 214);
            this.clstDBCubeName.Sorted = true;
            this.clstDBCubeName.TabIndex = 3;
            // 
            // cmdGenerateDocument
            // 
            this.cmdGenerateDocument.FlatAppearance.BorderSize = 5;
            this.cmdGenerateDocument.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdGenerateDocument.ForeColor = System.Drawing.Color.Blue;
            this.cmdGenerateDocument.Location = new System.Drawing.Point(178, 305);
            this.cmdGenerateDocument.Name = "cmdGenerateDocument";
            this.cmdGenerateDocument.Size = new System.Drawing.Size(331, 29);
            this.cmdGenerateDocument.TabIndex = 6;
            this.cmdGenerateDocument.Text = "Generate Document";
            this.cmdGenerateDocument.UseVisualStyleBackColor = true;
            this.cmdGenerateDocument.Click += new System.EventHandler(this.cmdGenerateDocument_Click);
            // 
            // txtProgress
            // 
            this.txtProgress.Location = new System.Drawing.Point(0, 353);
            this.txtProgress.Multiline = true;
            this.txtProgress.Name = "txtProgress";
            this.txtProgress.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtProgress.Size = new System.Drawing.Size(741, 158);
            this.txtProgress.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "DB * Cube Name  :";
            // 
            // cmdSelectAll
            // 
            this.cmdSelectAll.Location = new System.Drawing.Point(602, 48);
            this.cmdSelectAll.Name = "cmdSelectAll";
            this.cmdSelectAll.Size = new System.Drawing.Size(129, 23);
            this.cmdSelectAll.TabIndex = 4;
            this.cmdSelectAll.Text = "Select All";
            this.cmdSelectAll.UseVisualStyleBackColor = true;
            this.cmdSelectAll.Click += new System.EventHandler(this.cmdSelectAll_Click);
            // 
            // cmdUnselectAll
            // 
            this.cmdUnselectAll.Location = new System.Drawing.Point(603, 78);
            this.cmdUnselectAll.Name = "cmdUnselectAll";
            this.cmdUnselectAll.Size = new System.Drawing.Size(128, 23);
            this.cmdUnselectAll.TabIndex = 5;
            this.cmdUnselectAll.Text = "UnSelect All";
            this.cmdUnselectAll.UseVisualStyleBackColor = true;
            this.cmdUnselectAll.Click += new System.EventHandler(this.cmdUnselectAll_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 337);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Progress :";
            // 
            // txtOutputPath
            // 
            this.txtOutputPath.Location = new System.Drawing.Point(112, 268);
            this.txtOutputPath.Name = "txtOutputPath";
            this.txtOutputPath.Size = new System.Drawing.Size(484, 20);
            this.txtOutputPath.TabIndex = 3;
            // 
            // lblOutputPath
            // 
            this.lblOutputPath.AutoSize = true;
            this.lblOutputPath.Location = new System.Drawing.Point(8, 271);
            this.lblOutputPath.Name = "lblOutputPath";
            this.lblOutputPath.Size = new System.Drawing.Size(73, 13);
            this.lblOutputPath.TabIndex = 10;
            this.lblOutputPath.Text = "Output Path  :";
            // 
            // groupBox1
            // 
            this.groupBox1.Location = new System.Drawing.Point(0, 2);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(10);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(741, 296);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            // 
            // frmBISMDocumentor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(743, 517);
            this.Controls.Add(this.txtOutputPath);
            this.Controls.Add(this.lblOutputPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmdUnselectAll);
            this.Controls.Add(this.cmdSelectAll);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtProgress);
            this.Controls.Add(this.cmdGenerateDocument);
            this.Controls.Add(this.clstDBCubeName);
            this.Controls.Add(this.cmdConnect);
            this.Controls.Add(this.txtServerName);
            this.Controls.Add(this.lblServerName);
            this.Controls.Add(this.groupBox1);
            this.Name = "frmBISMDocumentor";
            this.Tag = " ";
            this.Text = "BISM - SSAS Tabular Documenter";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblServerName;
        private System.Windows.Forms.TextBox txtServerName;
        private System.Windows.Forms.Button cmdConnect;
        private System.Windows.Forms.CheckedListBox clstDBCubeName;
        private System.Windows.Forms.Button cmdGenerateDocument;
        private System.Windows.Forms.TextBox txtProgress;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button cmdSelectAll;
        private System.Windows.Forms.Button cmdUnselectAll;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtOutputPath;
        private System.Windows.Forms.Label lblOutputPath;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}