using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace excellDataReconstructor
{
    class ExcelToCsv
    {
        private Excel.Application application;
        private Excel.Workbook Workbook;

        //delete this
        private Excel.Worksheet newMySheet;
        private string currentSheet = "Sheet1";

        //delete this

        private string errorMessageBase = "Please select a destination to save the new excel document!";

        public string OrigionalFileUrl { get; set; }
        public string NewFileUrl { get; set; }

        public void ConvertToCsv()
        {
            if (NewFileUrl == null)
            {
                MessageBox.Show(errorMessageBase);
            }
            else
            {
                ConvertAndSave();
                Quit();
            }
        }

        private void ConvertAndSave()
        {
            application = new Excel.Application { Visible = false };
            Workbook = application.Workbooks.Open(OrigionalFileUrl);
            Workbook.WebOptions.Encoding = MsoEncoding.msoEncodingUTF8;
            Workbook.SaveAs(NewFileUrl, Excel.XlFileFormat.xlCSV);
        }

        private void Quit()
        {
            if (Workbook != null)
            {
                Workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            }
            if (application != null)
            {
                application.Quit();
            }
        }
    }
}
