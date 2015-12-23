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
            this.SuspendLayout();
            // 
            // openFileButton
            // 
            this.openFileButton.Location = new System.Drawing.Point(12, 12);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(186, 45);
            this.openFileButton.TabIndex = 0;
            this.openFileButton.Text = "Select origional Excel file";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.ReadFile_Click);
            // 
            // SaveFile
            // 
            this.SaveFile.Location = new System.Drawing.Point(13, 64);
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
            this.selectedExelNameLabel.Location = new System.Drawing.Point(204, 28);
            this.selectedExelNameLabel.Name = "selectedExelNameLabel";
            this.selectedExelNameLabel.Size = new System.Drawing.Size(95, 13);
            this.selectedExelNameLabel.TabIndex = 3;
            this.selectedExelNameLabel.Text = "Excel file selected:";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(437, 120);
            this.Controls.Add(this.selectedExelNameLabel);
            this.Controls.Add(this.SaveFile);
            this.Controls.Add(this.openFileButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "MainForm";
            this.Text = "Excel Parser";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.Button SaveFile;
        private System.Windows.Forms.Label selectedExelNameLabel;
    }
}

