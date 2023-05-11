using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ChecklistReview
{
    class AlternateExcel
    {
        public static void BuildWorkSheet(System.Data.DataTable table, List<string> specs)
        {
            Application xlApp = new Application();

            if(xlApp == null)
            {
                Console.WriteLine("Excel could not be started. Check if your office references are correct.");
                return;
            }
            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Add();
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            if (ws == null)
            { Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct."); }

            int i = 5;
            int c = 1;

            //insert column heads based on selected columns
            //loop through list of requested specs, set up the columns for that spec
            #region set up Columns
            //create ranges as columns
            foreach (string specName in specs)
            {
                Range dynamic;
                switch(specName)
                {   
                    // 1 column set to 11.5 width
                    case "StyleID":
                        Range StyleID = ws.Cells[2, c];
                        StyleID.Value = "Style ID";
                        StyleID.Font.Size = 10;
                        StyleID.RowHeight = 15;
                        StyleID.ColumnWidth = 11.45;
                        StyleID.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        Range lowerRow = ws.Cells[3, c];
                        lowerRow.RowHeight = 48.75;
                        c++;
                        break;

                    //1 Column | 8.25 width
                    case "Poly":
                        Range Poly = ws.Cells[2, c];
                        Poly.Value = "Poly";
                        Poly.Font.Size = 10;
                        Poly.ColumnWidth = 8.25;
                        Poly.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        break;

                    //1 Columns | 8.25 width
                    //3 Columns | 2.75 Width rotated vertical
                    case "Poly Quotes":
                        //set up upper part
                        Range PolyQuotes = ws.Range[ws.Cells[2, c], ws.Cells[2,(c+3)]];
                        PolyQuotes.Merge();
                        PolyQuotes.Value = "Poly Quotes";
                        PolyQuotes.Font.Size = 10;
                        PolyQuotes.HorizontalAlignment = XlHAlign.xlHAlignCenter;                        
                        //set up lower parts
                        Range quote = ws.Cells[3, c];
                        quote.Font.Size = 8;
                        quote.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        quote.VerticalAlignment = XlVAlign.xlVAlignBottom;
                        quote.ColumnWidth = 8.25;
                        quote.Value = "Quotes";
                        c++;
                        Range Mquote = ws.Cells[3, c];
                        Mquote.Font.Size = 8;
                        Mquote.Orientation = 90;
                        Mquote.ColumnWidth = 2.75;
                        Mquote.Value = "Marx";
                        c++;
                        Range Cquote = ws.Cells[3, c];
                        Cquote.Font.Size = 8;
                        Cquote.Orientation = 90;
                        Cquote.ColumnWidth = 2.75;
                        Cquote.Value = "Contour";
                        c++;
                        Range Squote = ws.Cells[3, c];
                        Squote.Font.Size = 8;
                        Squote.Orientation = 90;
                        Squote.ColumnWidth = 2.75;
                        Squote.Value = "Snyder";
                        c++;
                        break;

                    //1 Column | 8.25 w
                    case "Cardboard":
                        Range Cardboard = ws.Cells[2, c];
                        Cardboard.Value = "Cardboard";
                        Cardboard.Font.Size = 10;
                        Cardboard.ColumnWidth = 8.25;
                        Cardboard.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        break;

                    //1 Column | 8.25 w
                    case "Frame":
                        Range Frame = ws.Cells[2, c];
                        Frame.Value = "Frame";
                        Frame.Font.Size = 10;
                        Frame.ColumnWidth = 8.25;
                        Frame.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        break;

                    //1 Column | 8.25 w
                    case "Pattern":
                        Range Pattern = ws.Cells[2, c];
                        Pattern.Value = "Pattern";
                        Pattern.Font.Size = 10;
                        Pattern.ColumnWidth = 8.25;
                        Pattern.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        break;

                    //3 Columns | 8.25 w
                    case "Layouts":
                        Range Layouts = ws.Range[ws.Cells[2, c], ws.Cells[2, (c + 2)]];
                        Layouts.Merge();
                        Layouts.Value = "Layouts";
                        Layouts.Font.Size = 10;
                        Layouts.HorizontalAlignment = XlHAlign.xlHAlignCenter;                        
                        Range CR = ws.Cells[3, c];
                        CR.ColumnWidth = 8.25;
                        CR.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        CR.Font.Size = 8;
                        CR.Value = "CR";
                        c++;
                        Range RR = ws.Cells[3, c];
                        RR.ColumnWidth = 8.25;
                        RR.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        RR.Font.Size = 8;
                        RR.Value = "RR";
                        c++;
                        Range Plastic = ws.Cells[3, c];
                        Plastic.ColumnWidth = 8.25;
                        Plastic.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        Plastic.Font.Size = 8;
                        Plastic.Value = "Plastic";
                        c++;
                        break;

                    //1 Column | 8.25 w
                    case "Sewing":
                        Range Sewing = ws.Cells[2, c];
                        Sewing.Value = "Sewing";
                        Sewing.Font.Size = 10;
                        Sewing.ColumnWidth = 8.25;
                        Sewing.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        break;

                    //2 Columns | 8.25 w
                    case "Upholstery":
                        Range Uphol = ws.Range[ws.Cells[2, c], ws.Cells[2, (c + 1)]];
                        Uphol.Merge();
                        Uphol.Value = "Upholstery";
                        Uphol.Font.Size = 10;                        
                        Uphol.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        Range sketch = ws.Cells[3, c];
                        sketch.Value = "Sketch";
                        sketch.Font.Size = 8;
                        sketch.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        Range form = ws.Cells[3, c];
                        form.Value = "Form";
                        form.Font.Size = 8;
                        form.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        break;

                    //1 Column | 11.15 w
                    case "ProductInfo":
                        Range ProductInfo = ws.Cells[2, c];
                        ProductInfo.Value = "Product Info";
                        ProductInfo.Font.Size = 10;
                        ProductInfo.ColumnWidth = 11.15;
                        ProductInfo.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        break;

                    //1 Column |2.75 w rotated vertical
                    case "Dimensions":
                        Range dimensions = ws.Range[ws.Cells[2, c],ws.Cells[3,c]];
                        dimensions.Merge();
                        dimensions.Value = "Dimensions";
                        dimensions.Font.Size = 10;
                        dimensions.Orientation = 90;                        
                        dimensions.VerticalAlignment = XlVAlign.xlVAlignBottom;
                        dimensions.ColumnWidth = 2.75;
                        c++;
                        break;

                    //2 Columns | 3 w rotated vertical
                    case "Weights":
                        Range weights = ws.Range[ws.Cells[2, c], ws.Cells[2, (c + 1)]];
                        weights.Merge();
                        weights.Value = "Weights";                        
                        weights.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        weights.Font.Size = 10;
                        Range packaged = ws.Cells[3, c];
                        packaged.Value = "Packaged";
                        packaged.Orientation = 90;
                        packaged.VerticalAlignment = XlVAlign.xlVAlignBottom;
                        packaged.Font.Size = 8;
                        packaged.ColumnWidth = 3;
                        c++;
                        Range unpackaged = ws.Cells[3, c];
                        unpackaged.Value = "Unpackaged";
                        unpackaged.Orientation = 90;
                        unpackaged.VerticalAlignment = XlVAlign.xlVAlignBottom;
                        unpackaged.Font.Size = 8;
                        unpackaged.ColumnWidth = 3;
                        c++;
                        break;

                    //1 Column | 8.25
                    case "Cartoning":
                        Range Cartoning = ws.Cells[2, c];
                        Cartoning.Value = "Cartoning";
                        Cartoning.Font.Size = 10;
                        Cartoning.ColumnWidth = 8.25;
                        Cartoning.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        break;

                    //1 Column | 8.25
                    case "Photos":
                        Range Photos = ws.Cells[2, c];
                        Photos.Value = "Photos";
                        Photos.Font.Size = 10;
                        Photos.ColumnWidth = 8.25;
                        Photos.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        c++;
                        break;

                    default:
                        dynamic = ws.Cells[2, c];
                        dynamic.Font.Size = 10;
                        dynamic.ColumnWidth = 8;
                        dynamic.Value = specName;
                        dynamic.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        break;
                }
            }
            #endregion

            //based on specs chosen put date in the upper right corner
            //merge cells if the spot is too small.
            //the remaining space to the left is used for title
            #region dateBox            
            int lastcell = ws.UsedRange.Columns.Count;
            Range nextCell = ws.Cells[1, lastcell];
            double width = nextCell.ColumnWidth;
            while (width <= 8.25)
            {
                lastcell--;
                nextCell = ws.Cells[1, lastcell];
                width = width + nextCell.ColumnWidth;                
            }
            Range dateBox = ws.Range[ws.Cells[1, lastcell], ws.Cells[1, ws.UsedRange.Columns.Count]];
            dateBox.Merge();
            dateBox.Value = DateTime.Now.ToShortDateString();
            #endregion
            #region Date 
            lastcell--;
            int labelEnd = lastcell;          
            nextCell = ws.Cells[1, lastcell];
            width = nextCell.ColumnWidth;
            while (width <= 8.25)
            {
                lastcell--;
                nextCell = ws.Cells[1, lastcell];
                width = width + nextCell.ColumnWidth;
            }
            Range datelabel = ws.Range[ws.Cells[1, lastcell], ws.Cells[1, labelEnd]];
            datelabel.Merge();
            datelabel.Value = "Date:";
            datelabel.HorizontalAlignment = XlHAlign.xlHAlignRight;
            #endregion
            #region Title
            int titleEnd = lastcell - 1;
            Range title = ws.Range[ws.Cells[1, 1], ws.Cells[1, titleEnd]];
            title.Merge();
            title.RowHeight = 18;
            title.Font.Size = 14;
            title.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            title.Value = "Specification Check List";
            #endregion

            //loop through the datatable and display data based on the selected specs
            #region fill out table
            foreach (DataRow row in table.Rows)
            {
                //loop through the row to discern what will go into the cells
                //reset the column count for each row
                c = 1;
                foreach (string specName in specs)
                {
                    switch (specName)
                    {
                        case "StyleID":
                        ws.Cells[i, c++] = row["StyleID"]; //styleID
                            break;

                        case "Poly":
                        ws.Cells[i, c++] = row["Poly"]; //Poly
                            break;

                        case "Layouts":
                        ws.Cells[i, c++] = row["Pattern"]; //Pattern
                        if (row["CRSigned"].ToString() == "")
                        { ws.Cells[i, c++] = row["CR"]; } //CR
                        else
                        { ws.Cells[i, c++] = row["CR"] + " " + ((char)0x2713).ToString(); } //CR Signed
                        if (row["RRSigned"].ToString() == "")
                        { ws.Cells[i, c++] = row["RR"]; } //RR
                        else
                        { ws.Cells[i, c++] = row["RR"] + " " + ((char)0x2713).ToString(); }//RR Signed
                        ws.Cells[i, c++] = row["Plastic"]; //list if the plastic version of the layout exists
                            break;

                        case "Upholstery":
                        ws.Cells[i, c++] = row["Upholstery Sketch"]; //upol dwg
                        ws.Cells[i, c++] = row["Upholstery Form"]; //Uphol Form
                            break;

                        case "Cardboard":
                        ws.Cells[i, c++] = row["Cardboard"]; //CB
                            break;

                        case "Sewing":
                        ws.Cells[i, c++] = row["Sewing"]; //Sewing
                            break;

                        case "Frame":
                        ws.Cells[i, c++] = row["Frame"]; //Frame
                            break;

                        case "Photos":
                        ws.Cells[i, c++] = row["Photos"]; //Photos
                            break;

                        case "ProductInfo":
                        ws.Cells[i, c++] = row["Product Info"];//product Info
                            break;

                        case "Cartoning":
                        ws.Cells[i, c++] = row["Cartoning"]; //Cartoning 
                            break;

                        case "Dimensions":
                            if (row["Dimensions"].ToString() == "True")
                            { ws.Cells[i, c++] = ((char)0x2713).ToString(); }//unicode for checkmark
                            else c++;
                            break;

                        case "Weights":
                            if (row["Weight"].ToString() == "True")
                            { ws.Cells[i, c++] = ((char)0x2713).ToString(); }//Unpackaged Weight 
                            else c++;
                            if (row["Packaged Weight"].ToString() == "True")
                            { ws.Cells[i, c++] = ((char)0x2713).ToString(); }//Packaged Weight
                            else c++;
                            break;

                        case "Poly Quotes":
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
                            break;
                    }            
                }                        
                i++;
            }
            #endregion

            //set up borders around stuff
            //format the alignments and borders of the filled in cells
            Range all = ws.UsedRange;
            Range specList = ws.Range[ws.Cells[5, 1], ws.Cells[i, all.Columns.Count]];
            Range styleIds = ws.Range[ws.Cells[5, 1], ws.Cells[i, 1]];
            Range otherSpecs = ws.Range[ws.Cells[5, 2], ws.Cells[i, all.Columns.Count]];

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

            //set up the header to be frozen for prints and scrolling
            Range header = ws.Range[ws.Cells[1, ws.UsedRange.Columns.Count], ws.Cells[3, ws.UsedRange.Columns.Count]];
            ws.Application.ActiveWindow.SplitRow = 3;
            ws.Application.ActiveWindow.FreezePanes = true;
            header.AutoFilter(1,
                Type.Missing,
                XlAutoFilterOperator.xlAnd,
                Type.Missing,
                true);

            //print title so that prints have same header
            ws.PageSetup.PrintTitleRows = "$1:$3";

            //page set up for printing
            ws.PageSetup.Orientation = XlPageOrientation.xlLandscape;

            ws.PageSetup.BottomMargin = .5;
            ws.PageSetup.TopMargin = .5;
            ws.PageSetup.LeftMargin = .5;
            ws.PageSetup.RightMargin = .5;

            ws.PageSetup.Zoom = false;
            ws.PageSetup.FitToPagesWide = 1;
            ws.PageSetup.FitToPagesTall = false;
        }
    }
}
