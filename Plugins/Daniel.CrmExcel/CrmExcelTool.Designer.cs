namespace Daniel.CrmExcel
{
    partial class CrmExcelTool
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnClose = new System.Windows.Forms.Button();
            this.btnWhoAmI = new System.Windows.Forms.Button();
            this.grpUpdate = new System.Windows.Forms.GroupBox();
            this.cboSolutionUpdate = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cmdSelectExcelForUpdate = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.cmdUpdateCrm = new System.Windows.Forms.Button();
            this.textFileNameUpdate = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cmdSelectPrefix = new System.Windows.Forms.Button();
            this.txtPrefix = new System.Windows.Forms.TextBox();
            this.grpExcel = new System.Windows.Forms.GroupBox();
            this.btnAddEntitiesToInclude = new System.Windows.Forms.Button();
            this.txtEntitiesToInclude = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txtExcelFile = new System.Windows.Forms.TextBox();
            this.cboSolution = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.chkUseSolutionXml = new System.Windows.Forms.CheckBox();
            this.btnRefreshEntities = new System.Windows.Forms.Button();
            this.chkIncludeOwnerEtc = new System.Windows.Forms.CheckBox();
            this.cmdGenerateExcelSheet = new System.Windows.Forms.Button();
            this.tvwEntities = new System.Windows.Forms.TreeView();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.cmdRetieveCrmInformation = new System.Windows.Forms.Button();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.cboLanguage = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.grpUpdate.SuspendLayout();
            this.grpExcel.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(830, 3);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(52, 37);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.BtnCloseClick);
            // 
            // btnWhoAmI
            // 
            this.btnWhoAmI.Location = new System.Drawing.Point(749, 3);
            this.btnWhoAmI.Name = "btnWhoAmI";
            this.btnWhoAmI.Size = new System.Drawing.Size(75, 23);
            this.btnWhoAmI.TabIndex = 1;
            this.btnWhoAmI.Text = "Who Am I";
            this.btnWhoAmI.UseVisualStyleBackColor = true;
            this.btnWhoAmI.Click += new System.EventHandler(this.BtnWhoAmIClick);
            // 
            // grpUpdate
            // 
            this.grpUpdate.Controls.Add(this.cboSolutionUpdate);
            this.grpUpdate.Controls.Add(this.label6);
            this.grpUpdate.Controls.Add(this.cmdSelectExcelForUpdate);
            this.grpUpdate.Controls.Add(this.label4);
            this.grpUpdate.Controls.Add(this.cmdUpdateCrm);
            this.grpUpdate.Controls.Add(this.textFileNameUpdate);
            this.grpUpdate.Controls.Add(this.label5);
            this.grpUpdate.Location = new System.Drawing.Point(579, 61);
            this.grpUpdate.Name = "grpUpdate";
            this.grpUpdate.Size = new System.Drawing.Size(298, 343);
            this.grpUpdate.TabIndex = 11;
            this.grpUpdate.TabStop = false;
            this.grpUpdate.Text = "Update CRM";
            // 
            // cboSolutionUpdate
            // 
            this.cboSolutionUpdate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSolutionUpdate.FormattingEnabled = true;
            this.cboSolutionUpdate.Location = new System.Drawing.Point(12, 89);
            this.cboSolutionUpdate.Name = "cboSolutionUpdate";
            this.cboSolutionUpdate.Size = new System.Drawing.Size(280, 21);
            this.cboSolutionUpdate.TabIndex = 21;
            this.cboSolutionUpdate.SelectedIndexChanged += new System.EventHandler(this.cboSolutionUpdate_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 71);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 13);
            this.label6.TabIndex = 18;
            this.label6.Text = "Solution:";
            // 
            // cmdSelectExcelForUpdate
            // 
            this.cmdSelectExcelForUpdate.Location = new System.Drawing.Point(260, 129);
            this.cmdSelectExcelForUpdate.Name = "cmdSelectExcelForUpdate";
            this.cmdSelectExcelForUpdate.Size = new System.Drawing.Size(32, 20);
            this.cmdSelectExcelForUpdate.TabIndex = 17;
            this.cmdSelectExcelForUpdate.Text = "...";
            this.cmdSelectExcelForUpdate.UseVisualStyleBackColor = true;
            this.cmdSelectExcelForUpdate.Click += new System.EventHandler(this.cmdSelectExcelForUpdate_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(135, 52);
            this.label4.TabIndex = 14;
            this.label4.Text = "How to:\r\n1 Retrieve Metadata\r\n2 Selected edited Excel file\r\n3 Update CRM";
            // 
            // cmdUpdateCrm
            // 
            this.cmdUpdateCrm.Location = new System.Drawing.Point(42, 169);
            this.cmdUpdateCrm.Name = "cmdUpdateCrm";
            this.cmdUpdateCrm.Size = new System.Drawing.Size(203, 63);
            this.cmdUpdateCrm.TabIndex = 3;
            this.cmdUpdateCrm.Text = "Update CRM";
            this.cmdUpdateCrm.UseVisualStyleBackColor = true;
            this.cmdUpdateCrm.Click += new System.EventHandler(this.cmdUpdateCrm_Click);
            // 
            // textFileNameUpdate
            // 
            this.textFileNameUpdate.Location = new System.Drawing.Point(12, 129);
            this.textFileNameUpdate.Name = "textFileNameUpdate";
            this.textFileNameUpdate.Size = new System.Drawing.Size(242, 20);
            this.textFileNameUpdate.TabIndex = 16;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 113);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(88, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "Excel to process:";
            // 
            // cmdSelectPrefix
            // 
            this.cmdSelectPrefix.Location = new System.Drawing.Point(290, 190);
            this.cmdSelectPrefix.Name = "cmdSelectPrefix";
            this.cmdSelectPrefix.Size = new System.Drawing.Size(146, 20);
            this.cmdSelectPrefix.TabIndex = 7;
            this.cmdSelectPrefix.Text = "Check entities with prefix";
            this.cmdSelectPrefix.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.cmdSelectPrefix.UseVisualStyleBackColor = true;
            this.cmdSelectPrefix.Click += new System.EventHandler(this.cmdSelectPrefix_Click);
            // 
            // txtPrefix
            // 
            this.txtPrefix.Location = new System.Drawing.Point(236, 190);
            this.txtPrefix.Name = "txtPrefix";
            this.txtPrefix.Size = new System.Drawing.Size(48, 20);
            this.txtPrefix.TabIndex = 6;
            // 
            // grpExcel
            // 
            this.grpExcel.Controls.Add(this.btnAddEntitiesToInclude);
            this.grpExcel.Controls.Add(this.txtEntitiesToInclude);
            this.grpExcel.Controls.Add(this.button2);
            this.grpExcel.Controls.Add(this.button1);
            this.grpExcel.Controls.Add(this.btnSelectFile);
            this.grpExcel.Controls.Add(this.label3);
            this.grpExcel.Controls.Add(this.txtExcelFile);
            this.grpExcel.Controls.Add(this.cboSolution);
            this.grpExcel.Controls.Add(this.label1);
            this.grpExcel.Controls.Add(this.txtPrefix);
            this.grpExcel.Controls.Add(this.label2);
            this.grpExcel.Controls.Add(this.cmdSelectPrefix);
            this.grpExcel.Controls.Add(this.chkUseSolutionXml);
            this.grpExcel.Controls.Add(this.btnRefreshEntities);
            this.grpExcel.Controls.Add(this.chkIncludeOwnerEtc);
            this.grpExcel.Controls.Add(this.cmdGenerateExcelSheet);
            this.grpExcel.Controls.Add(this.tvwEntities);
            this.grpExcel.Location = new System.Drawing.Point(6, 61);
            this.grpExcel.Name = "grpExcel";
            this.grpExcel.Size = new System.Drawing.Size(567, 370);
            this.grpExcel.TabIndex = 10;
            this.grpExcel.TabStop = false;
            this.grpExcel.Text = "Excel Sheet generation";
            // 
            // btnAddEntitiesToInclude
            // 
            this.btnAddEntitiesToInclude.Location = new System.Drawing.Point(430, 333);
            this.btnAddEntitiesToInclude.Margin = new System.Windows.Forms.Padding(2);
            this.btnAddEntitiesToInclude.Name = "btnAddEntitiesToInclude";
            this.btnAddEntitiesToInclude.Size = new System.Drawing.Size(111, 25);
            this.btnAddEntitiesToInclude.TabIndex = 24;
            this.btnAddEntitiesToInclude.Text = "Add Entities To Include";
            this.btnAddEntitiesToInclude.UseVisualStyleBackColor = true;
            this.btnAddEntitiesToInclude.Click += new System.EventHandler(this.BtnAddEntitiesToInclude_Click);
            // 
            // txtEntitiesToInclude
            // 
            this.txtEntitiesToInclude.Location = new System.Drawing.Point(236, 338);
            this.txtEntitiesToInclude.Name = "txtEntitiesToInclude";
            this.txtEntitiesToInclude.Size = new System.Drawing.Size(189, 20);
            this.txtEntitiesToInclude.TabIndex = 23;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(441, 308);
            this.button2.Margin = new System.Windows.Forms.Padding(2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(111, 25);
            this.button2.TabIndex = 22;
            this.button2.Text = "WinCare";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(441, 259);
            this.button1.Margin = new System.Windows.Forms.Padding(2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(111, 37);
            this.button1.TabIndex = 21;
            this.button1.Text = "Fix Relations";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(524, 234);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(28, 20);
            this.btnSelectFile.TabIndex = 9;
            this.btnSelectFile.Text = "...";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(233, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(208, 78);
            this.label3.TabIndex = 13;
            this.label3.Text = "How to:\r\n1 Retrieve Metadata\r\n2 Select solution (with the form information)\r\n3 Se" +
    "lect / Enter Excel file\r\n4 Select Entities to include\r\n5 Generate Excel Sheet";
            // 
            // txtExcelFile
            // 
            this.txtExcelFile.Location = new System.Drawing.Point(236, 234);
            this.txtExcelFile.Name = "txtExcelFile";
            this.txtExcelFile.Size = new System.Drawing.Size(282, 20);
            this.txtExcelFile.TabIndex = 8;
            // 
            // cboSolution
            // 
            this.cboSolution.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSolution.FormattingEnabled = true;
            this.cboSolution.Location = new System.Drawing.Point(290, 128);
            this.cboSolution.Name = "cboSolution";
            this.cboSolution.Size = new System.Drawing.Size(262, 21);
            this.cboSolution.TabIndex = 20;
            this.cboSolution.SelectedIndexChanged += new System.EventHandler(this.CboSolutionSelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(235, 218);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Excel file to generate:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(233, 133);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Solution:";
            // 
            // chkUseSolutionXml
            // 
            this.chkUseSolutionXml.AutoSize = true;
            this.chkUseSolutionXml.Checked = true;
            this.chkUseSolutionXml.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUseSolutionXml.Location = new System.Drawing.Point(236, 105);
            this.chkUseSolutionXml.Name = "chkUseSolutionXml";
            this.chkUseSolutionXml.Size = new System.Drawing.Size(178, 17);
            this.chkUseSolutionXml.TabIndex = 5;
            this.chkUseSolutionXml.Text = "Include form location information";
            this.chkUseSolutionXml.UseVisualStyleBackColor = true;
            // 
            // btnRefreshEntities
            // 
            this.btnRefreshEntities.Location = new System.Drawing.Point(9, 19);
            this.btnRefreshEntities.Name = "btnRefreshEntities";
            this.btnRefreshEntities.Size = new System.Drawing.Size(94, 24);
            this.btnRefreshEntities.TabIndex = 4;
            this.btnRefreshEntities.Text = "Refresh Entities";
            this.btnRefreshEntities.UseVisualStyleBackColor = true;
            this.btnRefreshEntities.Click += new System.EventHandler(this.btnRefreshEntities_Click);
            // 
            // chkIncludeOwnerEtc
            // 
            this.chkIncludeOwnerEtc.AutoSize = true;
            this.chkIncludeOwnerEtc.Location = new System.Drawing.Point(236, 155);
            this.chkIncludeOwnerEtc.Name = "chkIncludeOwnerEtc";
            this.chkIncludeOwnerEtc.Size = new System.Drawing.Size(111, 17);
            this.chkIncludeOwnerEtc.TabIndex = 3;
            this.chkIncludeOwnerEtc.Text = "Include owner etc";
            this.chkIncludeOwnerEtc.UseVisualStyleBackColor = true;
            // 
            // cmdGenerateExcelSheet
            // 
            this.cmdGenerateExcelSheet.Location = new System.Drawing.Point(236, 275);
            this.cmdGenerateExcelSheet.Name = "cmdGenerateExcelSheet";
            this.cmdGenerateExcelSheet.Size = new System.Drawing.Size(189, 50);
            this.cmdGenerateExcelSheet.TabIndex = 0;
            this.cmdGenerateExcelSheet.Text = "Generate Excel";
            this.cmdGenerateExcelSheet.UseVisualStyleBackColor = true;
            this.cmdGenerateExcelSheet.Click += new System.EventHandler(this.cmdGenerateExcelSheet_Click);
            // 
            // tvwEntities
            // 
            this.tvwEntities.CheckBoxes = true;
            this.tvwEntities.FullRowSelect = true;
            this.tvwEntities.Location = new System.Drawing.Point(9, 49);
            this.tvwEntities.Name = "tvwEntities";
            this.tvwEntities.Size = new System.Drawing.Size(221, 276);
            this.tvwEntities.TabIndex = 2;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // cmdRetieveCrmInformation
            // 
            this.cmdRetieveCrmInformation.Location = new System.Drawing.Point(6, 18);
            this.cmdRetieveCrmInformation.Name = "cmdRetieveCrmInformation";
            this.cmdRetieveCrmInformation.Size = new System.Drawing.Size(451, 37);
            this.cmdRetieveCrmInformation.TabIndex = 12;
            this.cmdRetieveCrmInformation.Text = "Step 1 - Retrieve Metadata";
            this.cmdRetieveCrmInformation.UseVisualStyleBackColor = true;
            this.cmdRetieveCrmInformation.Click += new System.EventHandler(this.cmdRetieveCrmInformation_Click);
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog1";
            // 
            // cboLanguage
            // 
            this.cboLanguage.FormattingEnabled = true;
            this.cboLanguage.Items.AddRange(new object[] {
            "1043",
            "1033"});
            this.cboLanguage.Location = new System.Drawing.Point(554, 34);
            this.cboLanguage.Name = "cboLanguage";
            this.cboLanguage.Size = new System.Drawing.Size(68, 21);
            this.cboLanguage.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(463, 37);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(85, 13);
            this.label7.TabIndex = 21;
            this.label7.Text = "Language code:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(638, 37);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(101, 13);
            this.label8.TabIndex = 22;
            this.label8.Text = "1043 NL 1033 ENG";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(606, 418);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(18, 13);
            this.label9.TabIndex = 23;
            this.label9.Text = "x1";
            // 
            // CrmExcelTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cboLanguage);
            this.Controls.Add(this.cmdRetieveCrmInformation);
            this.Controls.Add(this.grpUpdate);
            this.Controls.Add(this.grpExcel);
            this.Controls.Add(this.btnWhoAmI);
            this.Controls.Add(this.btnClose);
            this.Name = "CrmExcelTool";
            this.Size = new System.Drawing.Size(885, 500);
            this.Load += new System.EventHandler(this.CrmExcelTool_Load);
            this.grpUpdate.ResumeLayout(false);
            this.grpUpdate.PerformLayout();
            this.grpExcel.ResumeLayout(false);
            this.grpExcel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnWhoAmI;
        private System.Windows.Forms.GroupBox grpUpdate;
        private System.Windows.Forms.Button cmdSelectPrefix;
        private System.Windows.Forms.TextBox txtPrefix;
        private System.Windows.Forms.Button cmdUpdateCrm;
        private System.Windows.Forms.GroupBox grpExcel;
        private System.Windows.Forms.CheckBox chkUseSolutionXml;
        private System.Windows.Forms.Button btnRefreshEntities;
        private System.Windows.Forms.CheckBox chkIncludeOwnerEtc;
        private System.Windows.Forms.Button cmdGenerateExcelSheet;
        private System.Windows.Forms.TreeView tvwEntities;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtExcelFile;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboSolution;
        private System.Windows.Forms.Button cmdRetieveCrmInformation;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textFileNameUpdate;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button cmdSelectExcelForUpdate;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.ComboBox cboSolutionUpdate;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cboLanguage;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnAddEntitiesToInclude;
        private System.Windows.Forms.TextBox txtEntitiesToInclude;
        private System.Windows.Forms.Label label9;
    }
}
