using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace excellDataReconstructor
{
    class ExcelParser
    {
        private Workbook originalWorkbookMyBook;
        private Application OriginalMyApp;
        private Worksheet OriginalMySheet;
      
        private Workbook newWorkbook;
        private Application newMyApp;
        private Worksheet newMySheet;

        private string currentSheet = "Sheet1";
        private string errorMessageBase = "Please select a destination to save the new excel document!";

        private List<List<string>> ContentMatrix = new List<List<string>>();
        List<string> column1StraingData = new List<string>();
        List<string> column2StraingData = new List<string>();

        public string OrigionalFileUrl { get; set; }
        public string NewFileUrl { get; set; }
        public int NumberOfRowsInOrigionalFile { get; set; }
        public int CurrentRowInFile { get; set; }

        public void Parse()
        {
            if (NewFileUrl == null)
            {
                MessageBox.Show(errorMessageBase);
            }
            else
            {
                Init();
                AddHeaders();
                GetColumnData();
                ParseData();
                AddContentToExcel();
                SaveData();
                MessageBox.Show("done");
                ClearDataForNextRun();
                Quit();
            }
        }

        private void Init()
        {
            //init origional excel file
            OriginalMyApp = new Application { Visible = false };
            originalWorkbookMyBook = OriginalMyApp.Workbooks.Open(OrigionalFileUrl);
            OriginalMySheet = (Worksheet)originalWorkbookMyBook.Sheets[1];

            //init new excel file
            newMyApp = new Application {Visible = false};
            newWorkbook = newMyApp.Workbooks.Add();
            newMySheet = (Worksheet)newWorkbook.Worksheets[1];
        }

        private void GetLastRow()
        {
            var Column = OriginalMySheet.Range["A:A"].Cells;
            foreach (Range item in Column.Cells)
            {
                if (item.Value != null && item.Value.ToString() == "Complete")
                {
                    break;
                }
                NumberOfRowsInOrigionalFile++;
            }
        }

        private void AddHeaders()
        {
            string[] firstRowAsArray =
            {
            "Company name", 
            "address", 
            "city", 
            "zip",
            "telephone 1",
            "telephone 2",
            "fax", 
            "email", 
            "cellular",
            "URL", 
            "I.D.",
            "ID activity",
            "number of employees",
            "contact1 first name", 
            "contact1 last name", 
            "contact1 title",
            "contact2 first name", 
            "contact2 last name", 
            "contact2 title"
        };
            string[] SecondRowAsArray =
        {
            "שם חברה", 
            "כתובת", 
            "ישוב", 
            "מיקוד",
            "טלפון 1",
            "טלפון 2",
            "פקס", 
            "דואר אלקטרוני", 
            "נייד",
            "כתובת אינטרנט", 
            "ח.פ./ע.מ.",
            "פעילות ה-ח.פ.",
            "מס. מועסקים",
            "איש קשר1 שם פרטי", 
            "איש קשר1 שם משפחה", 
            "איש קשר1 תפקיד",
            "איש קשר2 שם פרטי", 
            "איש קשר2 שם משפחה", 
            "איש קשר2 תפקיד"
        };
            ContentMatrix.Add(firstRowAsArray.ToList());
            ContentMatrix.Add(SecondRowAsArray.ToList());
        }

        private void GetColumnData()
        {
            var Column1 = OriginalMySheet.Range["A:A"].Cells.Value;
            var Column2 = OriginalMySheet.Range["B:B"].Cells.Value;
            foreach (var cellVal in Column1)
            {
                if (cellVal != null)
                {
                    if (cellVal.ToString() == "Complete")
                    {
                        break;
                    }
                    column1StraingData.Add(cellVal.ToString());
                }
                else
                {
                    column1StraingData.Add(null);
                }
            }
            foreach (var cellVal in Column2)
            {
                if (cellVal != null)
                {
                    if (cellVal.ToString() == "Complete")
                    {
                        break;
                    }

                    column2StraingData.Add(cellVal.ToString());
                }
                else
                {
                    column2StraingData.Add(null);
                }
            }
        }

        private void ParseData()
        {
            int i = 0;
            bool flag = false;
            List<string> row = new List<string>();

            foreach (var cellValueInCollumn2 in column2StraingData)
            {
                if (cellValueInCollumn2 != null)
                {
                    flag = true;
                    row = insertNextCell(cellValueInCollumn2, row);
                }
                else
                {
                    if (flag)
                    {
                        flag = false;
                        ContentMatrix.Add(row);
                        row = new List<string>();
                    }
                }

                i++;
            }
            ContentMatrix.Add(row);
        }

        private List<string> insertNextCell(string currentCell, List<string> currentRow)
        {
            int positionInArray = currentRow.Count;
            currentRow.Add(currentCell);
            return currentRow;
        }

        private void ClearDataForNextRun()
        {
            ContentMatrix = new List<List<string>>();
        }

        private void AddContentToExcel()
        {
            bool flag = true;
            while (flag)
            {
                try
                {
                    int i = 0;
                    int j = 0;
                    foreach (var cellList in ContentMatrix)
                    {
                        i++;
                        foreach (var contentCell in cellList)
                        {
                            j++;
                            newMySheet.Cells[i, j] = contentCell;
                        }
                        j = 0;
                    }
                    flag = false;
                }
                catch (Exception)
                {
                    // ignored
                }
            }
        }

        private void SaveData()
        {
            bool flag = true;
            while (flag)
            {
                try
                {
                    if (NewFileUrl == null)
                    {
                        MessageBox.Show(errorMessageBase);
                    }
                    else
                    {
                        newWorkbook.SaveAs(NewFileUrl);
                    }
                    flag = false;
                }
                catch (Exception)
                {
                    // ignored
                }
            }
        }

        private void Quit()
        {
            bool flag = true;
            while (flag)
            {
                try
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
                    flag = false;
                }
                catch (Exception)
                {
                    // ignored
                }
            }
        }
    }
}
