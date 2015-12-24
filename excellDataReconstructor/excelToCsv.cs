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
                Convert();
                Quit();
            }
        }

        private void Convert()
        {
            application = new Excel.Application { Visible = false };
            application.DefaultWebOptions.Encoding = MsoEncoding.msoEncodingUTF8;
            Workbook = application.Workbooks.Open(OrigionalFileUrl);

            Workbook.SaveAs(NewFileUrl);
        }

        private void Quit()
        {
            if (application != null)
            {
                application.Quit();
            }
        }
    }
}
