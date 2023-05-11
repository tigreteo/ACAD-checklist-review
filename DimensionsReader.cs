using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChecklistReview
{
    class DimensionsReader
    {
        //read through the Styles & Dim file to see if the styleid is listed and if it's weights are listed
        public static List<bool> readDimensions(string styleId, Worksheet ws)
        {
            Range all = ws.UsedRange;
            Range target = all.Find(styleId);

            List<bool> dimsExist = new List<bool>();

            if(target != null)
            {
                //first bool is for wether or not the style is in the list for styles and Dim
                dimsExist.Add(true);
                target = target.EntireRow;
                System.Array dimensions = (Array)target.Cells.Value;

                //if it contains a weight at the cell for unpackaged weight
                if(dimensions.GetValue(1,15) != null)
                { dimsExist.Add(true); }
                else
                { dimsExist.Add(false); }

                //if it contains a weight at the cell for packaged weight
                if (dimensions.GetValue(1, 16) != null)
                { dimsExist.Add(true); }
                else
                { dimsExist.Add(false); }
            }
            else
            {
            dimsExist.Add(false);
            dimsExist.Add(false);
            dimsExist.Add(false);
            }
            return dimsExist;
        }
    }
}
