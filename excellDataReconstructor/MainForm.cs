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
        DataReconstructor _DataReconstructor = new DataReconstructor();
        private const string ErrorMessageread = "Please select an origional excel file to parse!";
        private const string ErrorMessagesaveExcel = "Please select a destination to save the new excel document!";
        private const string ErrorMessagesaveCsv = "Please select a destination to save the new CSV document!";

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
            if (ParseDirectlyToCSV.Checked)
            {
                UpdateExcelltoCsvOriginalFile(_excelParser.NewFileUrl);
                string csvAdressUrl = _excelParser.NewFileUrl.Replace(".xlsx", ".csv").Replace(".xls", ".csv");
                UpdateNewExcelltoCsvOriginalFile(csvAdressUrl);
                ConvertToCsv();
                File.Delete(_excelParser.NewFileUrl);
            }
        }

        private void Parse()
        {
            if (_fileWizard.OrigionalFileUrl == null)
            {
                MessageBox.Show(ErrorMessageread);
            }
            else if (_fileWizard.NewFileUrl == null)
            {
                MessageBox.Show(ErrorMessagesaveExcel);
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
        private void UpdateExcelltoCsvOriginalFile(string url)
        {
            _csvConverter.OrigionalFileUrl = url;
        }
        private void UpdateNewExcelltoCsvOriginalFile()
        {
            _csvConverter.NewFileUrl = _fileWizard.NewCsvUrl;
        }
        private void UpdateNewExcelltoCsvOriginalFile(string Url)
        {
            _csvConverter.NewFileUrl = Url;
        }

        private void ConvertToCsv()
        {
            if (_csvConverter.OrigionalFileUrl == null)
            {
                MessageBox.Show(ErrorMessageread);
            }
            else if (_csvConverter.NewFileUrl == null)
            {
                MessageBox.Show(ErrorMessagesaveCsv);
            }
            else
            {
                _csvConverter.ConvertToCsv();
            }
        }

        private void SaveExcelFileFeildReconstructor_Click(object sender, EventArgs e)
        {
            //todo: _fileWizard.SelectSaveFileExcelDataReconstructor();
            //todo: add checks to see if the files are null
            if (true)
            {
                int i = (int) NumericRowSelectorExcelFileFeildReconstructor.Value;
                _DataReconstructor.Reconstruct(_fileWizard.OrigionalFileUrlDataReconstructor, _fileWizard.NewFileUrlDataReconstructor, _fileWizard.referenceFileUrlDataReconstructor, (int)NumericRowSelectorExcelFileFeildReconstructor.Value);
            }
        }

        private void SelectExcelFileFeildReconstructor_Click(object sender, EventArgs e)
        {
            _fileWizard.SelectFileToExcelDataReconstructor();
        }

        private void SelectExcelReferenceFileFeildReconstructor_Click(object sender, EventArgs e)
        {
            _fileWizard.SelectReferenceFileToExcelDataReconstructor();
        }
    }
}
