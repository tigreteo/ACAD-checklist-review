using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ChecklistReview
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    
    ///use form to determine if new excel sheet or just adding
    ///folder dialogue to find style(s)
    ///loop through styles on first through a LINQ search of the folder as well as for a .csv file of RR,CR layouts
    ///identify what type of spec the file referances
    ///determine if the type of file extension
    ///determine if the effect the type of file would have on that type of spec
    ///check if there is already data in the cell for that file 
    /// if there is, then compare the two to find the more relevant (ie date)
    ///use datatable to create an excel file
    public partial class App : System.Windows.Application
    {
        public static System.Data.DataTable Start(List<string> searchFiles = null)
        {
            //Styles and Dim file location
            // @"Y:\Customer Documents\Style Info\Styles and Dimensions.XLS"
            //CSV file location vvvvvvvvvvvvvvv
            string csvFile = @"Y:\Product Development\Standards\Layouts List.csv";
            IEnumerable<FileInfo> fileQuery = null;

            //VVv The below code creates a list of all fileInfos of anything they may qualify as a spec inside the search folder(s)
            //this is an inherantly ineffectiant design

            //if no list of locations to search exists, then default to hardcoded master file location
            if (searchFiles == null)
            { fileQuery = LINQ(@"Y:\Product Development\Style Specifications"); }
            else
            {
                //otherwise loop through folder locations
                foreach (string searchMe in searchFiles)
                {
                    //get enumeration of all files of particular extensions
                    IEnumerable<FileInfo> fi = LINQ(searchMe);
                    if (fileQuery == null)
                    { fileQuery = fi; }
                    else
                    { fileQuery = fileQuery.Concat(fi); }
                }
            }

            //!!!OLD CODE MOVED INTERFACE TO THE FIRST BOX
            //ask for the files to look through
            //request for a folder dialog
            //Used in testing
            //var dialog = new FolderBrowserDialog();
            //if (dialog.ShowDialog() == DialogResult.OK)
            //{ searchMe = dialog.SelectedPath; }                

            //set up a table to hold data as we process it
            System.Data.DataTable table = createTableHeader();

            //loop through each file in query and sort accordingly
            foreach (FileInfo fi in fileQuery)
            {
                //sort out what type of spec we're dealing with
                //then determine if that is the type of file it is 
                //depending on type of file and modification date, alter the cell data
                string file = fi.FullName;
                sortFiles(file, table);
            }
            
            //method to read CSV File
            //by running this second it will automatically be able to overwrite as higher priority
            if (File.Exists(csvFile))
            { table = readCSVFile(csvFile, table); }

            //check through all of the styleIDs in the datatable so far to see if it has dimensions and weights
            checkStyleAndDim(table);

            return table;
        }

        private static System.Data.DataTable readCSVFile(string csvFile, System.Data.DataTable table)
        {
            //read the csv file into an array, each line is an entry of a file
            string[] fileArray = File.ReadAllLines(csvFile);
            string[] currentRow;
            string[] lastStyle = { "-", "-", "-", "-", "-", };
            string[] dr = { "-", "-", "-", "-", "-", };
            foreach (string row in fileArray)
            {
                currentRow = row.Split(',');
                DataRow[] foundStyleId = table.Select("StyleID = '" + currentRow[0] + "'");
                //if it isnt already in the list then don't add it!!!!!!!!!
                //because the csv list is larger than the search parameters
                if (foundStyleId.Length == 0)
                {
                    //add the styleID to the table
                    //DataRow newRow = table.NewRow();
                    //newRow["StyleId"] = currentRow[0];
                    //add the spec data as well, since it must be new
                    //newRow[currentRow[1]] = currentRow[3];
                    //table.Rows.Add(newRow);
                }
                else
                {
                    //load the first (and hopefully only) data row from array
                    DataRow updateRow = foundStyleId[0];
                    //check if the row has data already or not
                    string currentCell = currentRow[1].ToString();
                    if (currentCell == "RR" || currentCell == "CR" || currentCell == "Plastic")
                    {
                        currentCell = updateRow[currentCell].ToString();
                        if (currentCell != "")
                        {
                            //compare the values
                            if (Convert.ToDateTime(currentCell) < Convert.ToDateTime(currentRow[3]))
                            { updateRow[currentRow[1]] = currentRow[3]; }
                        }
                        else
                        { updateRow[currentRow[1]] = currentRow[3]; }

                        //update the signedRow coloumn if it is listed as "Signed"
                        if (currentRow[1].ToString() == "RR" && currentRow[2].ToString() == "Signed")
                        { updateRow["RRSigned"] = "Signed"; }
                        if (currentRow[1].ToString() == "CR" && currentRow[2].ToString() == "Signed")
                        { updateRow["CRSigned"] = "Signed"; }
                    }
                }
            }
            return table;
        }

        //sort through the files to first determin what type of spec they refer to
        private static void sortFiles(string file, System.Data.DataTable table)
        {
            string specType = GetSpecFromFolder(file);
            string ext = Path.GetExtension(file);
            switch(specType)
            {
                case "ERROR NO SPECS":
                    //just write to styleID column
                    DataRow newRow = table.NewRow();
                    newRow["StyleId"] = GetStyleIDFromFolder(file);
                    table.Rows.Add(newRow);
                    break;
                    //if poly is a CAD drawing then it is valid
                    //if poly is only a PDF from vendor then its good to know, but less relevant
                    //if it is a xls then it is probably the poly compare chart, need to confirm and process it
                case "Poly":
                    if(ext == ".dwg" || ext == ".pdf")
                    { tableWriter(file, "Poly", table); }
                    if(ext == ".xls")
                    { polyCompare(file, table); }
                    break;
                    //the only pattern we care about is DWG
                case "Fabric":
                case "Pattern":
                    if (ext == ".dwg")
                    { tableWriter(file, "Pattern", table); }
                    break;
                    //if the uphol is a DWG then its a sketch
                    //if the uphol is an xls then its an uphol spec
                    //  in case of xls will probably need to pass the file to a reader to verify if there is a
                    //  insiding and outside spec
                case "Upholstery":
                    if (ext == ".dwg")
                    { tableWriter(file, "Upholstery Sketch", table); }
                    else if (ext == ".xls")
                    { tableWriter(file, "Upholstery Form", table); }
                    break;
                    //CB should only be a DWG
                    //Cardboard has empty txt files with the name of No Cardboard
                case "Cardboard":
                    if (ext == ".dwg")
                    { tableWriter(file, "Cardboard", table); }
                    else if(ext == ".txt")
                    { tableWriter(file, "Cardboard", table, true); }
                    break;
                    //Sewing should only be a DWG *However only a PDF would be a complete sewing spec
                case "Sewing":
                    if (ext == ".dwg")
                    { tableWriter(file, "Sewing", table); }
                    break;
                    //Frame specs should just be PDFs FOR NOW
                case "Frame":
                case "Spring Up":
                case "Spring-Up":
                    if (ext == ".pdf")
                    { tableWriter(file, "Frame", table); }
                    break;
                    //photos could be jpg, png, tif, probably pdf
                    //I'll let anything through for now
                case "Photos":
                    { tableWriter(file, "Photos", table); }
                    break;
                    //Cartoning specs will probably be dwgs and PDFs
                case "Cartoning":
                    { tableWriter(file, "Cartoning", table); }
                    break;
                    //product info is an xls and a PDF depending on how it was digitized
                case "Product Info":
                case "Product Information":
                    { tableWriter(file, "Product Info", table); }
                    break;
            }
        }

        //search the datatable for the cell based on the spec linked and the styleID
        private static void tableWriter(string file, string specName, System.Data.DataTable table, bool NA = false)
        {
            string styleID = GetStyleIDFromFolder(file);
            string modDate = File.GetLastWriteTime(file).ToShortDateString();
            DataRow[] foundStyleId = table.Select("StyleID = '" + styleID + "'");
            if (foundStyleId.Length == 0)
            {
                //add the styleID to the table
                DataRow newRow = table.NewRow();
                newRow["StyleId"] = styleID;
                //add the spec data as well, since it must be new
                if (NA == false)
                { newRow[specName] = modDate; }
                else
                    newRow[specName] = "NA";
                table.Rows.Add(newRow);
            }
            else
            {
                //load the first (and hopefully only) data row from array
                DataRow updateRow = foundStyleId[0];
                //check if the row has data already or not
                string currentCell = updateRow[specName].ToString();
                if(currentCell != "" && currentCell.ToUpper() != "NA")
                {
                    //compare the values
                    if(Convert.ToDateTime(currentCell) < Convert.ToDateTime(modDate))
                    {
                        if (NA == false)
                        {updateRow[specName] = modDate;}
                        else
                        { updateRow[specName] = "NA"; }
                    }
                }
                else
                {
                    if(NA == false)
                        { updateRow[specName] = modDate; }
                        else
                        { updateRow[specName] = "NA"; }
                }
            }
        }

        //verify if this excel file is a poly comparison file, then update the table with these notes
        public static void polyCompare(string file, System.Data.DataTable table)
        {
            //check if this file is a poly compare table
            if(Path.GetFileName(file).Contains("Poly Comparison"))
            {
                //read through the table to find out what companies have given quotes
                //just send the filePath to a reading app that returns the bools

                string styleID = GetStyleIDFromFolder(file);
                string modDate = File.GetLastWriteTime(file).ToShortDateString();
                DataRow[] foundStyleId = table.Select("StyleID = '" + styleID + "'");

                List<string> validVendors = readPolyQuotes.ValidateQuotes(file);
                DataRow targetRow;

                if(foundStyleId.Length == 0)
                { targetRow = table.NewRow(); }
                else
                {  targetRow = foundStyleId[0]; }

                //add the data to the table
                targetRow["StyleId"] = styleID;
                targetRow["PolyCost"] = File.GetLastWriteTime(file).ToShortDateString();
                //set each vendor as empty be default
                targetRow["PCostMarx"] = false;
                targetRow["PCostContour"] = false;
                targetRow["PCostSnyder"] = false;

                //check to see if any vendors we're listed
                foreach (string vendor in validVendors)
                {
                    switch (vendor)
                    {
                        case "Marx":
                            targetRow["PCostMarx"] = true;
                            break;
                        case "Contour":
                            targetRow["PCostContour"] = true;
                            break;
                        case "Snyder":
                            targetRow["PCostSnyder"] = true;
                            break;
                    }
                }
            }
        }

        //set up columns for datatable
        public static System.Data.DataTable createTableHeader()
        {
            // Set up table header
            System.Data.DataTable table = new System.Data.DataTable("Fabric Layout Report");

            //list table column headers
            DataColumn[] cols ={
                                  new DataColumn ("StyleID",typeof(string)),
                                  new DataColumn ("Poly", typeof(string)),
                                  new DataColumn ("Pattern", typeof(string)),
                                  new DataColumn ("CR",typeof(string)),
                                  new DataColumn ("RR",typeof(string)),
                                  new DataColumn ("CRSigned",typeof(string)),
                                  new DataColumn ("RRSigned",typeof(string)),
                                  new DataColumn ("Upholstery Sketch",typeof(string)),
                                  new DataColumn ("Upholstery Form",typeof(string)),
                                  new DataColumn ("Cardboard",typeof(string)),
                                  new DataColumn ("Sewing",typeof(string)),
                                  new DataColumn ("Frame",typeof(string)),
                                  new DataColumn ("Photos",typeof(string)),
                                  new DataColumn ("Dimensions", typeof(Boolean)),
                                  new DataColumn ("Weight", typeof(Boolean)),
                                  new DataColumn ("Packaged Weight", typeof(Boolean)),
                                  new DataColumn ("PolyCost", typeof(string)),
                                  new DataColumn ("PCostMarx", typeof(Boolean)),
                                  new DataColumn ("PCostContour", typeof(Boolean)),
                                  new DataColumn ("PCostSnyder", typeof(Boolean)),
                                  new DataColumn ("Product Info", typeof(string)),
                                  new DataColumn ("Cartoning", typeof(string)),
                                  new DataColumn ("Plastic", typeof(string))
                              };
            //load column headers into datatable
            table.Columns.AddRange(cols);
            return table;
        }      

        //create a query of all files under the folder chosen
        public static IEnumerable<System.IO.FileInfo> LINQ(string searchMe)
        {
            //If this styleID doesnt have specs, it should list with a false path 
            //return the path as a fileInfo
            if(searchMe.Contains("ERROR NO SPECS"))
            {
                return new FileInfo[] { new FileInfo(searchMe) };
            }

            //capture the file system
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(searchMe);

            // This method assumes that the application has discovery permissions 
            // for all folders under the specified path.
            IEnumerable<System.IO.FileInfo> fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories);

            //Create the query
            //***only looking for Acad files and pdfs
            //file organised by directories
            IEnumerable<System.IO.FileInfo> fileQuery =
                from file in fileList
                where file.Extension == ".dwg" || file.Extension == ".pdf" || file.Extension == ".xls"
                || file.Extension == ".tif" || file.Extension == ".jpg" || file.Extension == ".txt"
                orderby file.DirectoryName
                select file;

            return fileQuery;
        }

        private static string GetStyleIDFromFolder(string doc)
        {
            string styleID = "";
            string pathName = Path.GetDirectoryName(doc);
            string[] styleIDparts = pathName.Split('\\');
            styleID = styleIDparts[styleIDparts.Length - 2];
            return styleID;
        }

        private static string GetSpecFromFolder(string doc)
        {
            string spec = "";
            string pathName = Path.GetDirectoryName(doc);
            string[] styleIDparts = pathName.Split('\\');
            spec = styleIDparts[styleIDparts.Length - 1];
            return spec;
        }

        private static string GetStyleIDFromFile(string doc)
        {
            doc = Path.GetFileNameWithoutExtension(doc);
            string[] specNames = {"Poly",
                                 "Fabric",
                                 "Pattern",
                                 "Upholstery",
                                 "Cardboard",
                                 "Sewing",
                                 "Photos",
                                 "Spring-Up",
                                 "Spring Up",
                                 "Frame"};
            foreach(string specName in specNames)
            {
                doc = doc.Replace(specName, "");
            }
            doc = doc.Trim();
            return doc;
        }

        //loop through the dataRows and find whether or not the style is on file
        private static void checkStyleAndDim(System.Data.DataTable table)
        {
            //open up the file so it only has to be loaded once, close when finished
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = xlApp.Workbooks.Open(@"Y:\Customers\Style Info\Styles and Dimensions.XLS", null, true);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            foreach(DataRow row in table.Rows)
            {
                //pass the opened excel file and the style we're searching for
                List<bool> dimsExist = DimensionsReader.readDimensions(row[0].ToString(), ws);                

                //set the table with the bools
                row["Dimensions"] = dimsExist[0];
                row["Weight"] = dimsExist[1];
                row["Packaged Weight"] = dimsExist[2];
            }

            wb.Close();
        }
    }
}
