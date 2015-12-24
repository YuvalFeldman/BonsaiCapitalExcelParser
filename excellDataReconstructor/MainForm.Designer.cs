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
            this.SuspendLayout();
            // 
            // openFileButton
            // 
            this.openFileButton.Location = new System.Drawing.Point(12, 57);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(186, 45);
            this.openFileButton.TabIndex = 0;
            this.openFileButton.Text = "Select origional Excel file";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.ReadFile_Click);
            // 
            // SaveFile
            // 
            this.SaveFile.Location = new System.Drawing.Point(12, 137);
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
            this.selectedExelNameLabel.Location = new System.Drawing.Point(12, 106);
            this.selectedExelNameLabel.Name = "selectedExelNameLabel";
            this.selectedExelNameLabel.Size = new System.Drawing.Size(95, 13);
            this.selectedExelNameLabel.TabIndex = 3;
            this.selectedExelNameLabel.Text = "Excel file selected:";
            // 
            // selectExcelFileToCSV
            // 
            this.selectExcelFileToCSV.Location = new System.Drawing.Point(238, 57);
            this.selectExcelFileToCSV.Name = "selectExcelFileToCSV";
            this.selectExcelFileToCSV.Size = new System.Drawing.Size(185, 45);
            this.selectExcelFileToCSV.TabIndex = 4;
            this.selectExcelFileToCSV.Text = "Select Excel file";
            this.selectExcelFileToCSV.UseVisualStyleBackColor = true;
            this.selectExcelFileToCSV.Click += new System.EventHandler(this.selectExcelFileToCSV_Click);
            // 
            // SaveAsToCSV
            // 
            this.SaveAsToCSV.Location = new System.Drawing.Point(238, 137);
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
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Parse excel file";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(236, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Convert Excel to CSV";
            // 
            // exceltocsvlabel
            // 
            this.exceltocsvlabel.AutoSize = true;
            this.exceltocsvlabel.Location = new System.Drawing.Point(236, 106);
            this.exceltocsvlabel.Name = "exceltocsvlabel";
            this.exceltocsvlabel.Size = new System.Drawing.Size(92, 13);
            this.exceltocsvlabel.TabIndex = 8;
            this.exceltocsvlabel.Text = "Excel file selected";
            // 
            // label4
            // 
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label4.Location = new System.Drawing.Point(218, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(2, 180);
            this.label4.TabIndex = 9;
            // 
            // ParseDirectlyToCSV
            // 
            this.ParseDirectlyToCSV.AutoSize = true;
            this.ParseDirectlyToCSV.Location = new System.Drawing.Point(15, 40);
            this.ParseDirectlyToCSV.Name = "ParseDirectlyToCSV";
            this.ParseDirectlyToCSV.Size = new System.Drawing.Size(125, 17);
            this.ParseDirectlyToCSV.TabIndex = 10;
            this.ParseDirectlyToCSV.Text = "Parse directly to CSV";
            this.ParseDirectlyToCSV.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(437, 196);
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
    }
}

