using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace excellDataReconstructor
{
    class ExcelParser
    {
        private Excel.Workbook originalWorkbookMyBook = null;
        private Excel.Application OriginalMyApp = null;
        private Excel.Worksheet OriginalMySheet = null;
      
        private Excel.Workbook newWorkbook;
        private Excel.Application newMyApp;
        private Excel.Sheets excelSheets;
        private Excel.Worksheet newMySheet;

        string currentSheet = "Sheet1";

        public string OrigionalFileUrl { get; set; }
        public string NewFileUrl { get; set; }

        public void Parse()
        {
            init();
            addHeaders();
            SaveData();
            Quit();
        }

        private void init()
        {
            //init origional excel file
            //OriginalMyApp = new Excel.Application();
            //OriginalMyApp.Visible = true;
            //originalWorkbookMyBook = OriginalMyApp.Workbooks.Open(OrigionalFileUrl);
            //OriginalMySheet = (Excel.Worksheet)originalWorkbookMyBook.Sheets[1];
            //var OriginallastRow = OriginalMySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            //init new excel file
            newMyApp = new Excel.Application {Visible = false};
            newWorkbook = newMyApp.Workbooks.Add();
            excelSheets = newWorkbook.Worksheets;
            newMySheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
        }

        private void addHeaders()
        {
            newMySheet.Cells[1, 1] = "test";
        }

        private void SaveData()
        {
            newWorkbook.SaveAs(NewFileUrl);
        }

        private void Quit()
        {
            newMyApp.Quit();
        }
    }
}
