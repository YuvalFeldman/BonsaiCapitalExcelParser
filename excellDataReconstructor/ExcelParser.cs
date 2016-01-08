using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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
        private const string errorMessageBase = "Please select a destination to save the new excel document!";
        private const string TaskCompletedMessage = "Conversion to excel has completed.";
        private const string endString = "מספר דנס";

        private List<List<string>> ContentMatrix = new List<List<string>>();
        List<string> ContentColumnData = new List<string>();
        List<string> HeaderColumnData = new List<string>();

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
                Init();
                AddHeaders();
                GetColumnData();
                FixColumnData();
                ParseData();
                AddContentToExcel();
                SaveData();
                ClearDataForNextRun();
                Quit();
                ShowTaskCompleteMessage();
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
            var column1 = OriginalMySheet.Range["A:A"].Cells.Value;
            var column2 = OriginalMySheet.Range["B:B"].Cells.Value;
            foreach (var cellVal in column1)
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
            foreach (var cellVal in column2)
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

        private void FixColumnData()
        {
            string[] headerOptions =
            {
                "שם חברה",
                "כתובת",
                "ישוב",
                "טלפון",
                "פקס",
                "דואר אלקטרוני",
                "כתובת אינטרנט",
                "מס. רישום",
                "סיווגים",
                "אופי פעילות",
                "מס. מועסקים",
                "מנהלים",
                "מספר דנס"
            };
            for(int k = 0; k < HeaderColumnData.Count; k++)
            {
                var flag = true;
                if (HeaderColumnData[k] == null) continue;
                if (headerOptions.Any(headerOption => HeaderColumnData[k] == headerOption))
                {
                    flag = false;
                }
                if (!flag) continue;
                var closestHeader = "";
                var maxAmountOfMatchingCharacters = 0;

                foreach (var headerOption in headerOptions)
                {
                    if (headerOption.Length != HeaderColumnData[k].Length) continue;
                    var headerOptionArray = headerOption.ToCharArray();
                    var headerColumnDataArray = HeaderColumnData[k].ToCharArray();
                    var numberOfSimilareChars = headerOptionArray.Where((t, i) => t == headerColumnDataArray[i]).Count();
                    if (numberOfSimilareChars <= maxAmountOfMatchingCharacters) continue;
                    closestHeader = headerOption;
                    maxAmountOfMatchingCharacters = numberOfSimilareChars;
                }

                HeaderColumnData[k] = closestHeader;
            }
        }

        private void ParseData()
        {
            int i = 0;
            List<string> headerRow = new List<string>();
            List<string> contentRow = new List<string>();

            foreach (var cellValueInCollumn2 in HeaderColumnData)
            {
                if (cellValueInCollumn2 != null)
                {
                    if (cellValueInCollumn2 == endString)
                    {
                        ContentMatrix.Add(FormatRow(headerRow, contentRow));
                        headerRow = new List<string>();
                        contentRow = new List<string>();
                    }

                    headerRow.Add(cellValueInCollumn2);
                    contentRow.Add(ContentColumnData[i]);
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
                string[] seperatedContentByComa;
                switch (unformatedHeaderRow[i])
                {
                    case "שם חברה":
                        formatedContentRow[0] = unformatedContentRow[i];
                        break;

                    case "כתובת":
                        formatedContentRow[1] = unformatedContentRow[i];
                        break;

                    case "ישוב":
                        seperatedContentByComa = unformatedContentRow[i].Split(',');
                        formatedContentRow[2] = seperatedContentByComa[0];

                        if (seperatedContentByComa.Length > 1 && seperatedContentByComa[1] != null)
                        {
                            formatedContentRow[3] = seperatedContentByComa[1];
                        }
                        break;

                    case "טלפון":
                        seperatedContentByComa = unformatedContentRow[i].Split(',');

                        formatedContentRow[4] = seperatedContentByComa[0];

                        if (seperatedContentByComa.Length == 2 && seperatedContentByComa[1] != null)
                        {
                            formatedContentRow[8] = seperatedContentByComa[1];
                        }
                        else if (seperatedContentByComa.Length > 3 && seperatedContentByComa[2] != null)
                        {
                            formatedContentRow[5] = seperatedContentByComa[1];
                            formatedContentRow[8] = seperatedContentByComa[2];
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
                        seperatedContentByComa = unformatedContentRow[i].Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        var furtherSeperatedContent = seperatedContentByComa[0].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        switch (furtherSeperatedContent.Length)
                        {
                            case 1:
                                formatedContentRow[13] = furtherSeperatedContent[0];
                                break;
                            default:
                                formatedContentRow[13] = furtherSeperatedContent[0];
                                string joinedContent = furtherSeperatedContent[1];
                                for (int l = 2; l < furtherSeperatedContent.Length; l++)
                                {
                                    joinedContent = string.Format("{0} {1}", joinedContent, furtherSeperatedContent[l]);
                                }
                                formatedContentRow[14] = joinedContent;
                                break;
                        }

                        if (seperatedContentByComa.Length >= 2 && seperatedContentByComa[1] != null)
                        {
                            furtherSeperatedContent = seperatedContentByComa[1].Split(new[]{' '}, StringSplitOptions.RemoveEmptyEntries);
                            switch (furtherSeperatedContent.Length)
                            {
                                case 1:
                                    formatedContentRow[15] = furtherSeperatedContent[0];
                                    break;
                                case 2:
                                    formatedContentRow[15] = furtherSeperatedContent[0];
                                    formatedContentRow[16] = furtherSeperatedContent[1];
                                    break;
                                default:
                                    formatedContentRow[15] = furtherSeperatedContent[0];
                                    formatedContentRow[16] = furtherSeperatedContent[1];

                                    string joinedContent = furtherSeperatedContent[2];
                                    for (int l = 3; l < furtherSeperatedContent.Length; l++)
                                    {
                                        joinedContent = string.Format("{0} {1}", joinedContent, furtherSeperatedContent[l]);
                                    }
                                    formatedContentRow[17] = joinedContent;
                                    break;
                            }
                        }
                        if (seperatedContentByComa.Length >= 3 && seperatedContentByComa[2] != null)
                        {
                            furtherSeperatedContent = seperatedContentByComa[2].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            formatedContentRow[18] = furtherSeperatedContent[0];

                        }
                        break;
                }
            }

            return formatedContentRow.ToList();
        }

        private void ClearDataForNextRun()
        {
            ContentMatrix = new List<List<string>>();
            ContentColumnData = new List<string>();
            HeaderColumnData = new List<string>();
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

        private void ShowTaskCompleteMessage()
        {
            MessageBox.Show(TaskCompletedMessage);
        }
    }
}
