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
        private SaveFileDialog _saveFileDialog = new SaveFileDialog();

        public FileWizard()
        {
            _openFileDialogDisplayContent.Filter = @"Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx| Excel Files(*.xlsm)|*.xlsm";
            _saveFileDialog.Filter = @"Excel Files(.xlsx)|*.xlsx| Excel Files(*.xlsm)|*.xlsm";
        }

        public string OrigionalFileUrl { get; set; }
        public string OrigionalFileName { get; set; }
        public string NewFileUrl { get; set; }

        public void SelectFile()
        {
            if (_openFileDialogDisplayContent.ShowDialog() == DialogResult.OK)
            {
                OrigionalFileUrl = _openFileDialogDisplayContent.FileName;
                OrigionalFileName = _openFileDialogDisplayContent.SafeFileName;
            }
        }

        public void CreateNewExcelFile()
        {
            if (_saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                NewFileUrl = _saveFileDialog.FileName;
                //File.Create(NewFileUrl);
            }
        }
    }
}
