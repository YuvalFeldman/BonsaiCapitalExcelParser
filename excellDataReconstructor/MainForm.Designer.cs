namespace excellDataReconstructor
{
    partial class MainForm
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
            this.openFileButton = new System.Windows.Forms.Button();
            this.SaveFile = new System.Windows.Forms.Button();
            this.selectedExelNameLabel = new System.Windows.Forms.Label();
            this.selectExcelFileToCSV = new System.Windows.Forms.Button();
            this.SaveAsToCSV = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.exceltocsvlabel = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.ParseDirectlyToCSV = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.SelectExcelFileFeildReconstructor = new System.Windows.Forms.Button();
            this.SaveExcelFileFeildReconstructor = new System.Windows.Forms.Button();
            this.NumericRowSelectorExcelFileFeildReconstructor = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.FileSelectedLabelExcelFileFeildReconstructor = new System.Windows.Forms.Label();
            this.SelectExcelReferenceFileFeildReconstructor = new System.Windows.Forms.Button();
            this.ExcelReferenceFileLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.NumericRowSelectorExcelFileFeildReconstructor)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileButton
            // 
            this.openFileButton.Location = new System.Drawing.Point(11, 48);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(186, 45);
            this.openFileButton.TabIndex = 0;
            this.openFileButton.Text = "Select origional Excel file";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.ReadFile_Click);
            // 
            // SaveFile
            // 
            this.SaveFile.Location = new System.Drawing.Point(11, 146);
            this.SaveFile.Name = "SaveFile";
            this.SaveFile.Size = new System.Drawing.Size(185, 45);
            this.SaveFile.TabIndex = 2;
            this.SaveFile.Text = "Save as...";
            this.SaveFile.UseVisualStyleBackColor = true;
            this.SaveFile.Click += new System.EventHandler(this.SaveFile_Click);
            // 
            // selectedExelNameLabel
            // 
            this.selectedExelNameLabel.AutoSize = true;
            this.selectedExelNameLabel.Location = new System.Drawing.Point(11, 97);
            this.selectedExelNameLabel.Name = "selectedExelNameLabel";
            this.selectedExelNameLabel.Size = new System.Drawing.Size(95, 13);
            this.selectedExelNameLabel.TabIndex = 3;
            this.selectedExelNameLabel.Text = "Excel file selected:";
            // 
            // selectExcelFileToCSV
            // 
            this.selectExcelFileToCSV.Location = new System.Drawing.Point(237, 48);
            this.selectExcelFileToCSV.Name = "selectExcelFileToCSV";
            this.selectExcelFileToCSV.Size = new System.Drawing.Size(185, 45);
            this.selectExcelFileToCSV.TabIndex = 4;
            this.selectExcelFileToCSV.Text = "Select Excel file";
            this.selectExcelFileToCSV.UseVisualStyleBackColor = true;
            this.selectExcelFileToCSV.Click += new System.EventHandler(this.selectExcelFileToCSV_Click);
            // 
            // SaveAsToCSV
            // 
            this.SaveAsToCSV.Location = new System.Drawing.Point(237, 146);
            this.SaveAsToCSV.Name = "SaveAsToCSV";
            this.SaveAsToCSV.Size = new System.Drawing.Size(186, 45);
            this.SaveAsToCSV.TabIndex = 5;
            this.SaveAsToCSV.Text = "Save as...";
            this.SaveAsToCSV.UseVisualStyleBackColor = true;
            this.SaveAsToCSV.Click += new System.EventHandler(this.SaveAsToCSV_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Parse excel file";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(234, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Convert Excel to CSV";
            // 
            // exceltocsvlabel
            // 
            this.exceltocsvlabel.AutoSize = true;
            this.exceltocsvlabel.Location = new System.Drawing.Point(235, 97);
            this.exceltocsvlabel.Name = "exceltocsvlabel";
            this.exceltocsvlabel.Size = new System.Drawing.Size(92, 13);
            this.exceltocsvlabel.TabIndex = 8;
            this.exceltocsvlabel.Text = "Excel file selected";
            // 
            // label4
            // 
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label4.Location = new System.Drawing.Point(217, 11);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(2, 180);
            this.label4.TabIndex = 9;
            // 
            // ParseDirectlyToCSV
            // 
            this.ParseDirectlyToCSV.AutoSize = true;
            this.ParseDirectlyToCSV.Location = new System.Drawing.Point(14, 31);
            this.ParseDirectlyToCSV.Name = "ParseDirectlyToCSV";
            this.ParseDirectlyToCSV.Size = new System.Drawing.Size(125, 17);
            this.ParseDirectlyToCSV.TabIndex = 10;
            this.ParseDirectlyToCSV.Text = "Parse directly to CSV";
            this.ParseDirectlyToCSV.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label3.Location = new System.Drawing.Point(446, 11);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(2, 180);
            this.label3.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(474, 11);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(99, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "Feild Reconstructor";
            // 
            // SelectExcelFileFeildReconstructor
            // 
            this.SelectExcelFileFeildReconstructor.Location = new System.Drawing.Point(477, 31);
            this.SelectExcelFileFeildReconstructor.Name = "SelectExcelFileFeildReconstructor";
            this.SelectExcelFileFeildReconstructor.Size = new System.Drawing.Size(96, 45);
            this.SelectExcelFileFeildReconstructor.TabIndex = 13;
            this.SelectExcelFileFeildReconstructor.Text = "Select Excel File to reconstruct";
            this.SelectExcelFileFeildReconstructor.UseVisualStyleBackColor = true;
            this.SelectExcelFileFeildReconstructor.Click += new System.EventHandler(this.SelectExcelFileFeildReconstructor_Click);
            // 
            // SaveExcelFileFeildReconstructor
            // 
            this.SaveExcelFileFeildReconstructor.Location = new System.Drawing.Point(477, 164);
            this.SaveExcelFileFeildReconstructor.Name = "SaveExcelFileFeildReconstructor";
            this.SaveExcelFileFeildReconstructor.Size = new System.Drawing.Size(204, 27);
            this.SaveExcelFileFeildReconstructor.TabIndex = 14;
            this.SaveExcelFileFeildReconstructor.Text = "Save as...";
            this.SaveExcelFileFeildReconstructor.UseVisualStyleBackColor = true;
            this.SaveExcelFileFeildReconstructor.Click += new System.EventHandler(this.SaveExcelFileFeildReconstructor_Click);
            // 
            // NumericRowSelectorExcelFileFeildReconstructor
            // 
            this.NumericRowSelectorExcelFileFeildReconstructor.Location = new System.Drawing.Point(625, 136);
            this.NumericRowSelectorExcelFileFeildReconstructor.Name = "NumericRowSelectorExcelFileFeildReconstructor";
            this.NumericRowSelectorExcelFileFeildReconstructor.Size = new System.Drawing.Size(56, 20);
            this.NumericRowSelectorExcelFileFeildReconstructor.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(474, 138);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(145, 13);
            this.label6.TabIndex = 16;
            this.label6.Text = "Select column to reconstruct ";
            // 
            // FileSelectedLabelExcelFileFeildReconstructor
            // 
            this.FileSelectedLabelExcelFileFeildReconstructor.AutoSize = true;
            this.FileSelectedLabelExcelFileFeildReconstructor.Location = new System.Drawing.Point(474, 85);
            this.FileSelectedLabelExcelFileFeildReconstructor.Name = "FileSelectedLabelExcelFileFeildReconstructor";
            this.FileSelectedLabelExcelFileFeildReconstructor.Size = new System.Drawing.Size(95, 13);
            this.FileSelectedLabelExcelFileFeildReconstructor.TabIndex = 17;
            this.FileSelectedLabelExcelFileFeildReconstructor.Text = "Excel file selected:";
            // 
            // SelectExcelReferenceFileFeildReconstructor
            // 
            this.SelectExcelReferenceFileFeildReconstructor.Location = new System.Drawing.Point(580, 31);
            this.SelectExcelReferenceFileFeildReconstructor.Name = "SelectExcelReferenceFileFeildReconstructor";
            this.SelectExcelReferenceFileFeildReconstructor.Size = new System.Drawing.Size(101, 45);
            this.SelectExcelReferenceFileFeildReconstructor.TabIndex = 18;
            this.SelectExcelReferenceFileFeildReconstructor.Text = "Select Excel reference file";
            this.SelectExcelReferenceFileFeildReconstructor.UseVisualStyleBackColor = true;
            this.SelectExcelReferenceFileFeildReconstructor.Click += new System.EventHandler(this.SelectExcelReferenceFileFeildReconstructor_Click);
            // 
            // ExcelReferenceFileLabel
            // 
            this.ExcelReferenceFileLabel.AutoSize = true;
            this.ExcelReferenceFileLabel.Location = new System.Drawing.Point(474, 111);
            this.ExcelReferenceFileLabel.Name = "ExcelReferenceFileLabel";
            this.ExcelReferenceFileLabel.Size = new System.Drawing.Size(143, 13);
            this.ExcelReferenceFileLabel.TabIndex = 19;
            this.ExcelReferenceFileLabel.Text = "Excel reference file selected:";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(697, 207);
            this.Controls.Add(this.ExcelReferenceFileLabel);
            this.Controls.Add(this.SelectExcelReferenceFileFeildReconstructor);
            this.Controls.Add(this.FileSelectedLabelExcelFileFeildReconstructor);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.NumericRowSelectorExcelFileFeildReconstructor);
            this.Controls.Add(this.SaveExcelFileFeildReconstructor);
            this.Controls.Add(this.SelectExcelFileFeildReconstructor);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ParseDirectlyToCSV);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.exceltocsvlabel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.SaveAsToCSV);
            this.Controls.Add(this.selectExcelFileToCSV);
            this.Controls.Add(this.selectedExelNameLabel);
            this.Controls.Add(this.SaveFile);
            this.Controls.Add(this.openFileButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "MainForm";
            this.Text = "Excel Parser";
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.NumericRowSelectorExcelFileFeildReconstructor)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.Button SaveFile;
        private System.Windows.Forms.Label selectedExelNameLabel;
        private System.Windows.Forms.Button selectExcelFileToCSV;
        private System.Windows.Forms.Button SaveAsToCSV;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label exceltocsvlabel;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox ParseDirectlyToCSV;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button SelectExcelFileFeildReconstructor;
        private System.Windows.Forms.Button SaveExcelFileFeildReconstructor;
        private System.Windows.Forms.NumericUpDown NumericRowSelectorExcelFileFeildReconstructor;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label FileSelectedLabelExcelFileFeildReconstructor;
        private System.Windows.Forms.Button SelectExcelReferenceFileFeildReconstructor;
        private System.Windows.Forms.Label ExcelReferenceFileLabel;
    }
}

