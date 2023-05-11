using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChecklistReview
{
    class ReadPreviousCheckList
    {
        //read through the file looking for styleIds in expected location (assuming format of checklist review)
        public static List<string> findStyleIds(string file)
        {
            List<string> styleIds = new List<string>();
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = xlApp.Workbooks.Open(file, null, true);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            Range rng = ws.UsedRange;

            //start on row 5 column A
            for(int i = 5; i< rng.Rows.Count; i++)
            {
                string cellValue = getCellValue(i, 1, rng);
                if(cellValue != null)
                {
                    styleIds.Add(cellValue);
                }
            }

            return styleIds;
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
