using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace excellDataReconstructor
{

    class FileWizard
    {
        private OpenFileDialog _openFileDialogDisplayContent = new OpenFileDialog();
        private SaveFileDialog _saveFileDialogExcelFilter = new SaveFileDialog();
        private SaveFileDialog _saveFileDialogCSVFilter = new SaveFileDialog();

        public FileWizard()
        {
            _openFileDialogDisplayContent.Filter = @"Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx| Excel Files(*.xlsm)|*.xlsm";
            _saveFileDialogExcelFilter.Filter = @"Excel Files(.xlsx)|*.xlsx| Excel Files(*.xlsm)|*.xlsm";
            _saveFileDialogCSVFilter.Filter = @"CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml";
        }

        public string OrigionalFileUrl { get; set; }
        public string OrigionalFileName { get; set; }
        public string NewFileUrl { get; set; }
        public string OrigionalExcelToCsvUrl { get; set; }
        public string OrigionalExcelToCsvFileName { get; set; }
        public string NewCsvUrl { get; set; }
        public void SelectFileToExcel()
        {
            OrigionalFileUrl = null;
            OrigionalFileName = null;
            if (_openFileDialogDisplayContent.ShowDialog() == DialogResult.OK)
            {
                OrigionalFileUrl = _openFileDialogDisplayContent.FileName;
                OrigionalFileName = _openFileDialogDisplayContent.SafeFileName;
            }
        }

        public void SelectFileToCsv()
        {
            OrigionalExcelToCsvUrl = null;
            OrigionalExcelToCsvFileName = null;
            if (_openFileDialogDisplayContent.ShowDialog() == DialogResult.OK)
            {
                OrigionalExcelToCsvUrl = _openFileDialogDisplayContent.FileName;
                OrigionalExcelToCsvFileName = _openFileDialogDisplayContent.SafeFileName;
            }
        }

        public void SelectSaveFileExcel()
        {
            NewFileUrl = null;
            if (_saveFileDialogExcelFilter.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(NewFileUrl) && NewFileUrl != null)
                {
                    File.Delete(NewFileUrl);
                }
                NewFileUrl = _saveFileDialogExcelFilter.FileName;
            }
        }

        public void SelectSaveFileCsv()
        {
            NewCsvUrl = null;
            if (_saveFileDialogCSVFilter.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(NewCsvUrl) && NewCsvUrl != null)
                {
                    File.Delete(NewCsvUrl);
                }
                NewCsvUrl = _saveFileDialogCSVFilter.FileName;
            }
        }
    }
}
