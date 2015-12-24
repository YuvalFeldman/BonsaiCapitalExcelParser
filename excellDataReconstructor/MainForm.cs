using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace excellDataReconstructor
{
    public partial class MainForm : Form
    {
        FileWizard _fileWizard = new FileWizard();
        ExcelParser _excelParser = new ExcelParser();

        public MainForm()
        {
            InitializeComponent();
        }

        private void ReadFile_Click(object sender, EventArgs e)
        {
            _fileWizard.SelectFile();
            UpdateOrigionalExcellParserFileUrl();
            UpdateExcelSelectedFileLabel();
        }

        private void SaveFile_Click(object sender, EventArgs e)
        {
            _fileWizard.CreateNewExcelFile();
            UpdateNewExcellParserFileUrl();
            _excelParser.Parse();
        }

        private void UpdateOrigionalExcellParserFileUrl()
        {
            _excelParser.OrigionalFileUrl = _fileWizard.OrigionalFileUrl;
        }

        private void UpdateNewExcellParserFileUrl()
        {
            _excelParser.NewFileUrl = _fileWizard.NewFileUrl;
        }

        private void UpdateExcelSelectedFileLabel()
        {
            selectedExelNameLabel.Text = string.Format("Excel file selected: {0}", _fileWizard.OrigionalFileName);
        }
    }
}
