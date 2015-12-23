using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace excellDataReconstructor
{
    class ExcelParser
    {
        public string OrigionalFileUrl { get; set; }
        public string NewFileUrl { get; set; }

        public void Parse()
        {
            addHeaders();
        }

        private void addHeaders()
        {
            
        }
    }
}
