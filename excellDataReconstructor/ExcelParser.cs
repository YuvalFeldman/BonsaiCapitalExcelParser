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
        List<string> ContentColumnData = new List<string>();
        List<string> HeaderColumnData = new List<string>();

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
                    ContentColumnData.Add(cellVal.ToString());
                }
                else
                {
                    ContentColumnData.Add(null);
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

                    HeaderColumnData.Add(cellVal.ToString());
                }
                else
                {
                    HeaderColumnData.Add(null);
                }
            }
        }

        private void ParseData()
        {
            int i = 0;
            bool flag = false;
            List<string> headerRow = new List<string>();
            List<string> contentRow = new List<string>();

            foreach (var cellValueInCollumn2 in HeaderColumnData)
            {
                if (cellValueInCollumn2 != null)
                {
                    flag = true;
                    headerRow.Add(cellValueInCollumn2);
                    contentRow.Add(ContentColumnData[i]);
                }
                else
                {
                    if (flag)
                    {
                        flag = false;
                        ContentMatrix.Add(FormatRow(headerRow, contentRow));
                        headerRow = new List<string>();
                        contentRow = new List<string>();
                    }
                }

                i++;
            }
            ContentMatrix.Add(headerRow);
        }

        private List<string> FormatRow(List<string> headerRow, List<string> contentRow)
        {
            int amountOfHeaders = ContentMatrix[1].Count;
            string[] formatedContentRow = new string[amountOfHeaders];
            string[] unformatedHeaderRow = headerRow.ToArray();
            string[] unformatedContentRow = contentRow.ToArray();

            for (int i = 0; i < amountOfHeaders; i++)
            {
                formatedContentRow[i] = null;
            }

            for (int i = 0; i < unformatedHeaderRow.Length; i++)
            {
                string[] seperatedContent;
                switch (unformatedHeaderRow[i])
                {
                    case "שם חברה":
                        formatedContentRow[0] = unformatedContentRow[i];
                        break;

                    case "כתובת":
                        formatedContentRow[1] = unformatedContentRow[i];
                        break;

                    case "ישוב":
                        seperatedContent = unformatedContentRow[i].Split(',');
                        formatedContentRow[2] = seperatedContent[0];

                        if (seperatedContent.Length > 1 && seperatedContent[1] != null)
                        {
                            formatedContentRow[3] = seperatedContent[1];
                        }
                        break;

                    case "טלפון":
                        seperatedContent = unformatedContentRow[i].Split(',');

                        formatedContentRow[4] = seperatedContent[0];

                        if (seperatedContent.Length > 1 && seperatedContent[1] != null)
                        {
                            formatedContentRow[5] = seperatedContent[1];
                        }

                        if (seperatedContent.Length > 2 && seperatedContent[2] != null)
                        {
                            formatedContentRow[8] = seperatedContent[2];
                        }

                        break;

                    case "פקס":
                        formatedContentRow[6] = unformatedContentRow[i];
                        break;

                    case "דואר אלקטרוני":
                        formatedContentRow[7] = unformatedContentRow[i];
                        break;

                    case "כתובת אינטרנט":
                        formatedContentRow[9] = unformatedContentRow[i];
                        break;

                    case "מס. רישום":
                        formatedContentRow[10] = unformatedContentRow[i];
                        break;

                    case "סיווגים":
                        if (formatedContentRow[11] == null)
                        {
                            formatedContentRow[11] = unformatedContentRow[i];
                        }
                        else
                        {
                            formatedContentRow[11] = string.Format("{0} {1}", unformatedContentRow[i], formatedContentRow[11]);
                        }
                        break;

                    case "אופי פעילות":
                        if (formatedContentRow[11] == null)
                        {
                            formatedContentRow[11] = unformatedContentRow[i];
                        }
                        else
                        {
                            formatedContentRow[11] = string.Format("{0} {1}", unformatedContentRow[i], formatedContentRow[11]);
                        }
                        break;

                    case "מס. מועסקים":
                        formatedContentRow[12] = unformatedContentRow[i];
                        break;

                    case "מנהלים":
                        seperatedContent = unformatedContentRow[i].Split(',');

                        var furtherSeperatedContent = seperatedContent[0].Split(' ');
                        switch (furtherSeperatedContent.Length)
                        {
                            case 1:
                                formatedContentRow[13] = furtherSeperatedContent[0];
                                break;
                            case 2:
                                formatedContentRow[13] = furtherSeperatedContent[0];
                                formatedContentRow[14] = furtherSeperatedContent[1];
                                break;
                            case 3:
                                formatedContentRow[13] = furtherSeperatedContent[0];
                                formatedContentRow[14] = string.Format("{0} {1}", furtherSeperatedContent[1], furtherSeperatedContent[2]);
                                break;
                        }

                        if (seperatedContent.Length > 1 && seperatedContent[1] != null)
                        {
                            formatedContentRow[15] = seperatedContent[1];
                        }

                        if (seperatedContent.Length > 2 && seperatedContent[2] != null)
                        {
                            furtherSeperatedContent = seperatedContent[2].Split(' ');

                            switch (furtherSeperatedContent.Length)
                            {
                                case 1:
                                    formatedContentRow[16] = furtherSeperatedContent[0];
                                    break;
                                case 2:
                                    formatedContentRow[16] = furtherSeperatedContent[0];
                                    formatedContentRow[17] = furtherSeperatedContent[1];
                                    break;
                                case 3:
                                    formatedContentRow[16] = furtherSeperatedContent[0];
                                    formatedContentRow[17] = string.Format("{0} {1}", furtherSeperatedContent[1], furtherSeperatedContent[2]);
                                    break;
                            }
                        }

                        if (seperatedContent.Length > 3 && seperatedContent[3] != null)
                        {
                            formatedContentRow[18] = seperatedContent[3];
                        }

                        break;
                }
            }

            return formatedContentRow.ToList();
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
