using System;
using System.Data;
using Microsoft.Office.Interop.Excel;

//fills out a pre-template of an excel sheet for specs

    //!!!!!!!!!!!!!REPLACED WITH A DYNAMIC EXCEL SHEET BUILDER!!!!!!!!!!!!
namespace ChecklistReview
{
    class BuildExcel
    {
        public static void BuildWorkSheet(System.Data.DataTable table, bool keepOldData = false)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if(xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Open(@"Y:\Product Development\Forms\Specification Check List.xls");
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            if (ws == null)
            { Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct."); }
            
            //Starting on row 5 to ignore the header in the pre-saved form if new
            int i = 5;
            //if the add button was clicked, then it will need to add to the end of the table
            //add is intended for when generating a report on multiple StyleIDs
            if(keepOldData == true)
            {
                Range rng = ws.UsedRange;
                i = rng.Rows.Count + 1;
            }
            //column counter
            int c = 1;
            //add the current date into the file
            //ws.Cells[1, ws.UsedRange.Columns.Count] = DateTime.Now.ToShortDateString();
            ws.Cells[1, 19] = DateTime.Now.ToShortDateString();

            #region Fill out data
            //loop through DataTable rows and format borders
            //by using an int for the columns instead of hardcoding we can move the columns wherever we want
            //should we make the shown columns something that is click in advance in the next build
            foreach (DataRow row in table.Rows)
            {
                c = 1;
                //loop through the row to discern what will go into the cells
                ws.Cells[i, c++] = row["StyleID"]; //styleID
                ws.Cells[i, c++] = row["Poly"]; //Poly
                ws.Cells[i, c++] = row["Pattern"]; //Pattern

                if (row["CRSigned"].ToString() == "")
                { ws.Cells[i, c++] = row["CR"]; } //CR
                else
                { ws.Cells[i, c++] = row["CR"] + " " +((char)0x2713).ToString(); } //CR Signed

                if (row["RRSigned"].ToString() == "")
                { ws.Cells[i, c++] = row["RR"]; } //RR
                else
                { ws.Cells[i, c++] = row["RR"] + " " +((char)0x2713).ToString(); }//RR Signed

                ws.Cells[i, c++] = row["Plastic"]; //list if the plastic version of the layout exists
                ws.Cells[i, c++] = row["Upholstery Sketch"]; //upol dwg
                ws.Cells[i, c++] = row["Upholstery Form"]; //Uphol Form
                ws.Cells[i, c++] = row["Cardboard"]; //CB
                ws.Cells[i, c++] = row["Sewing"]; //Sewing
                ws.Cells[i, c++] = row["Frame"]; //Frame
                ws.Cells[i, c++] = row["Photos"]; //Photos
                ws.Cells[i, c++] = row["Product Info"];//product Info
                ws.Cells[i, c++] = row["Cartoning"]; //Cartoning 

                if (row["Dimensions"].ToString() == "True")
                { ws.Cells[i, c++] = ((char)0x2713).ToString(); }//unicode for checkmark
                else c++;
                if (row["Weight"].ToString() == "True")
                { ws.Cells[i, c++] = ((char)0x2713).ToString(); }//Unpackaged Weight    
                else c++;
                if (row["Packaged Weight"].ToString() == "True")
                { ws.Cells[i, c++] = ((char)0x2713).ToString(); }//Packaged Weight
                else c++;

                ws.Cells[i, c++] = row["PolyCost"]; //PolyCost Date                
                if (row["PCostMarx"].ToString() == "True")
                { ws.Cells[i, c++] = ((char)0x2713).ToString(); }//marx
                else c++;
                if (row["PCostContour"].ToString() == "True")
                { ws.Cells[i, c++] = ((char)0x2713).ToString(); }//contour
                else c++;
                if (row["PCostSnyder"].ToString() == "True")
                { ws.Cells[i, c++] = ((char)0x2713).ToString(); }//snyder  
                else c++;                      
                i++;
            }

            #endregion

            //format the alignments and borders of the filled in cells
            Range all = ws.UsedRange;
            Range specList = ws.Range[ws.Cells[5,1], ws.Cells[i,all.Columns.Count]];            
            Range styleIds = ws.Range[ws.Cells[5,1], ws.Cells[i,1]];
            Range otherSpecs = ws.Range[ws.Cells[5,2], ws.Cells[i, all.Columns.Count]];

            //Keep the fonts small so that it can print
            specList.Font.Size = 8;
            styleIds.Font.Size = 10;

            //make the whole doc surrounded by think black lines
            Borders black = all.Borders;
            black[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            black[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            black[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            black[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            black.Weight = XlBorderWeight.xlThin;
            black.ColorIndex = 1;

            //make all of the basic data enveloped by thin grey lines
            Borders grey = specList.Borders;
            grey[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            grey[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            grey[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            grey[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            grey[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            grey[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            grey.Weight = XlBorderWeight.xlThin;
            grey.ColorIndex = 15;

            //align the style Ids to the left
            otherSpecs.HorizontalAlignment = XlHAlign.xlHAlignLeft;

            //align the other Specs centered
            otherSpecs.HorizontalAlignment = XlHAlign.xlHAlignCenter;


        }
    }
}
