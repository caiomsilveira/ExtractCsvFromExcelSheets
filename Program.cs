using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace ExtractCsvFromExcelSheets
{
    class Program
    {
        private static string CleanName(string name)
        {
            return name
                .Replace("<", "")
                .Replace(">", "")
                .Replace("?", "")
                .Replace("[", "")
                .Replace("]", "")
                .Replace(":", "")
                .Replace("|", "")
                .Replace("*", "");
        }

        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();

            string source = ConfigurationManager.AppSettings["SOURCE_FOLDER"].ToString();
            string target = ConfigurationManager.AppSettings["TARGET_FOLDER"].ToString();

            foreach (string file in System.IO.Directory.GetFiles(source, "*.xlsx"))
            {
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(file);
                try
                {
                    xlApp.Visible = false;
                    xlApp.DisplayAlerts = false;
                    foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
                    {

                        //var LastColumn = sheet.UsedRange.Columns.Count;
                        //LastColumn = LastColumn + sheet.UsedRange.Column - 1;
                        //int i = 0;
                        //for (i = 1; i <= LastColumn; i++)
                        //{
                        //    if (xlApp.WorksheetFunction.Count(sheet.Columns[i]) == 0)
                        //        (xlApp.Columns[i] as Microsoft.Office.Interop.Excel.Range).Delete();
                        //}
                        
                        sheet.Select();
                        xlWorkBook.SaveAs(string.Format("{0}{1}.csv", target, CleanName(sheet.Name)), Excel.XlFileFormat.xlCSV, Excel.XlSaveAsAccessMode.xlNoChange);

                    }
                }
                finally
                {
                    xlWorkBook.Close(false);
                }
            }

            Console.ReadKey();
        }
    }
}
