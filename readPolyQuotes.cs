using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace ChecklistReview
{
    class readPolyQuotes
    {
        //read through the file looking to see if there is any data for each company listed
        public static List<string> ValidateQuotes(string file)
        {
            List<string> quotes = new List<string>();
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = xlApp.Workbooks.Open(file, null, true);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            string[] vendors = new string[]{
                "Contour",
                "Marx",
                "Snyder"};

            Range rng = ws.UsedRange;

            for(int i = 1; i< rng.Rows.Count; i++)
            {
                string cellValue = getCellValue(i, 1, rng);
                if (cellValue != null)
                {
                    foreach (string vendor in vendors)
                    {
                        if (cellValue.Contains(vendor))
                        {
                            //check to see if anything is in the table
                            i = i + 2;
                            string nextCell = getCellValue(i, 1, rng);
                            if (nextCell != "" && nextCell != null)
                            {
                                string validVendor = vendor;
                                quotes.Add(validVendor);
                            }
                        }
                    }
                }
            }

            wb.Close(false);
            return quotes;
        }

        private static string getCellValue(int row, int column, Range rng)
        {
            var cellValue = "";
            try
            { cellValue = (string)(rng.Cells[row, column] as Range).Value; }
            catch (NullReferenceException)
            { cellValue = null; }

            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                var dt = (rng.Cells[row, column] as Range).Value;
                cellValue = dt.ToString();
            }

            return cellValue;
        }
    }
}
