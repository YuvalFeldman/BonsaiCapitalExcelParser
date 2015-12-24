using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
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
        private Excel.Worksheet newMySheet;

        private string currentSheet = "Sheet1";
        private string errorMessageBase = "Please select a destination to save the new excel document!";

        public string OrigionalFileUrl { get; set; }
        public string NewFileUrl { get; set; }

        public void Parse()
        {
            if (NewFileUrl == null)
            {
                MessageBox.Show(errorMessageBase);
            }
            else
            {
                init();
                addHeaders();
                SaveData();
                Quit();
            }
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
            newMySheet = (Excel.Worksheet)newWorkbook.Worksheets[1];
        }

        private void addHeaders()
        {
            newMySheet.Cells[1, 1] = "test";
        }

        private void SaveData()
        {
            if (NewFileUrl == null)
            {
                MessageBox.Show(errorMessageBase);
            }
            else
            {
                newWorkbook.SaveAs(NewFileUrl);
            }
        }

        private void Quit()
        {
            if (originalWorkbookMyBook != null)
            {
                originalWorkbookMyBook.Close();
            }
            if (newWorkbook != null)
            {
                newWorkbook.Close();
            }
            if (OriginalMyApp != null)
            {
                OriginalMyApp.Quit();
            }
            if (newMyApp != null)
            {
                newMyApp.Quit();
            }
        }
    }
}
