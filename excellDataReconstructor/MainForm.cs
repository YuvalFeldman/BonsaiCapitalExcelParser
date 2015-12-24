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
        ExcelToCsv _csvConverter = new ExcelToCsv();
        private string errorMessageread = "Please select an origional excel file to parse!";
        private string errorMessagesaveExcel = "Please select a destination to save the new excel document!";
        private string errorMessagesaveCsv = "Please select a destination to save the new CSV document!";

        public MainForm()
        {
            InitializeComponent();
        }

        private void ReadFile_Click(object sender, EventArgs e)
        {
            _fileWizard.SelectFileToExcel();
            UpdateOrigionalExcellParserFileUrl();
            UpdateExcelSelectedFileLabel();
        }

        private void SaveFile_Click(object sender, EventArgs e)
        {
            _fileWizard.SelectSaveFileExcel();
            UpdateNewExcellParserFileUrl();
            Parse();
        }

        private void Parse()
        {
            if (_fileWizard.OrigionalFileUrl == null)
            {
                MessageBox.Show(errorMessageread);
            }
            else if (_fileWizard.NewFileUrl == null)
            {
                MessageBox.Show(errorMessagesaveExcel);
            }
            else
            {
                _excelParser.Parse();
            }
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

        private void MainForm_Load(object sender, EventArgs e)
        {

        }
        private void selectExcelFileToCSV_Click(object sender, EventArgs e)
        {
            _fileWizard.SelectFileToCsv();
            UpdateNewExcelltoCsvFileLabel();
            UpdateExcelltoCsvOriginalFile();
        }

        private void SaveAsToCSV_Click(object sender, EventArgs e)
        {
            _fileWizard.SelectSaveFileCsv();
            UpdateNewExcelltoCsvOriginalFile();
            ConvertToCsv();
        }

        private void UpdateNewExcelltoCsvFileLabel()
        {
            exceltocsvlabel.Text = string.Format("Excel file selected: {0}", _fileWizard.OrigionalExcelToCsvFileName);
        }

        private void UpdateExcelltoCsvOriginalFile()
        {
            _csvConverter.OrigionalFileUrl = _fileWizard.OrigionalExcelToCsvUrl;
        }
        private void UpdateNewExcelltoCsvOriginalFile()
        {
            _csvConverter.NewFileUrl = _fileWizard.NewCsvUrl;
        }

        private void ConvertToCsv()
        {
            if (_fileWizard.OrigionalExcelToCsvUrl == null)
            {
                MessageBox.Show(errorMessageread);
            }
            else if (_fileWizard.NewCsvUrl == null)
            {
                MessageBox.Show(errorMessagesaveCsv);
            }
            else
            {
                _csvConverter.ConvertToCsv();
            }
        }
    }
}
