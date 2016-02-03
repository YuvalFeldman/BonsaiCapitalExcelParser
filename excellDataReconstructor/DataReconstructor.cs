using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace excellDataReconstructor
{
    class DataReconstructor
    {
        private string OrigionalFileUrl { get; set; }
        private string ReferenceFileUrl { get; set; }
        private string SaveFileUrl { get; set; }
        private int ColumnToReconstruct { get; set; }

        private List<string> CorrectDataList = new List<string>();
        private List<string> CoruptDataList = new List<string>();
        private List<string> FixedDataList = new List<string>();

        private Workbook originalWorkbookMyBook;
        private Application OriginalMyApp;
        private Worksheet OriginalMySheet;
        private Workbook referenceWorkbookMyBook;
        private Application referenceMyApp;
        private Worksheet referenceMySheet;

        public void Reconstruct(string origionalFileUrl, string newFileUrl, string referenceFileUrl, int columnToReconstruct)
        {
            OrigionalFileUrl = origionalFileUrl;
            ReferenceFileUrl = referenceFileUrl;
            SaveFileUrl = newFileUrl;
            ColumnToReconstruct = columnToReconstruct;

            Init();

            GetCorrectDataList();
            GetCoruptDataList();
        }

        private void Init()
        {
            //init origional excel file
            OriginalMyApp = new Application { Visible = false };
            originalWorkbookMyBook = OriginalMyApp.Workbooks.Open(OrigionalFileUrl);
            OriginalMySheet = (Worksheet)originalWorkbookMyBook.Sheets[1];

            //init origional excel file
            referenceMyApp = new Application { Visible = false };
            referenceWorkbookMyBook = referenceMyApp.Workbooks.Open(ReferenceFileUrl);
            referenceMySheet = (Worksheet)referenceWorkbookMyBook.Sheets[1];
        }

        private string getColumnString()
        {
            char columnLetter = (char)(ColumnToReconstruct + 65);
            return string.Format("{0}:{0}", columnLetter);
        }

        private void GetCorrectDataList()
        {
            var columnData = referenceMySheet.Range["A:A"].Cells.Value;
            foreach (var cellVal in columnData)
            {
                CorrectDataList.Add(cellVal != null ? cellVal.ToString() : null);
            }
        }
        private void GetCoruptDataList()
        {
            var columnData = OriginalMySheet.Range[getColumnString()].Cells.Value;
            foreach (var cellVal in columnData)
            {
                CoruptDataList.Add(cellVal != null ? cellVal.ToString() : null);
            }
        }

        private void constructFixedDataList()
        {
            
        }
    }
}
