using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Forms;
using System.IO;
using System;
using System.Linq;

namespace ChecklistReview
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<string> styleFolders = new List<string>();
        public List<string> foldernames = new List<string>();
        public IEnumerable<FileInfo> masterList = new DirectoryInfo(@"Y:\Product Development\Style Specifications").GetFiles("*.*", SearchOption.AllDirectories);

        //TODO
        //be able to add styleIDs that dont have specs in product & development

        //what happens when a styleID is in the list that doesnt have any specs?

        public MainWindow()
        {
            InitializeComponent();
        }

        //ask user for file to link to
        //pull data from file (in an expected format) and add StyleID numbers
        private void Import_Click(object sender, RoutedEventArgs e)
        {
            //get file location of excel file to read from
            var dialog = new OpenFileDialog();
            dialog.InitialDirectory = @"C:\Users\thenderson\Desktop";
            dialog.Filter = "excel files (*.xlsx)| *.xlsx|All files (*.*)| *.*";
            dialog.FilterIndex = 2;
            dialog.RestoreDirectory = true;
            //DialogResult result = dialog.ShowDialog();
            //dialog.ShowDialog();

            //load file and read list of style-IDs
            if (dialog.ShowDialog().ToString() == "OK")
            {
                List<string> styleIds = ReadPreviousCheckList.findStyleIds(dialog.FileName);

                //verify styleIDs
                foreach (string id in styleIds)
                {
                    //assume it is in @"Y:\Product Development\Style Specifications" but not testfolder
                    //use linq to find folders matching that description
                    string testFile = findFolder(@"Y:\Product Development\Style Specifications", id);

                    //add the folder to the list
                    if (Directory.Exists(testFile))
                    {
                        styleFolders.Add(testFile);
                        foldernames.Add(Path.GetFileName(testFile));
                    }
                    else
                    {
                        styleFolders.Add(id + @"\ERROR NO SPECS\ERROR.nope");
                        foldernames.Add(id);
                    }
                }
            }
            
            //update list and if list contains anything (after verification) then show go button
            if(foldernames.Count >0)
            {
                //display list to user
                displayList.ItemsSource = foldernames;
                displayList.Items.Refresh();

                if (GoButton.IsVisible == false)
                    GoButton.Visibility = Visibility.Visible;
            }           
        }

        //**TODO get the field to retain last position of folder selected
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //if the field for style ID has a style ID            
            string testString = insertStyleId.Text;
            testString = testString.Trim();
            if (testString == "")
            {
                //if the field for styleID was blank
                //request a folder to search through
                var dialog = new FolderBrowserDialog();
                dialog.SelectedPath = @"Y:\Product Development\Style Specifications";
                dialog.ShowDialog();

                //add the folder to the list
                styleFolders.Add(dialog.SelectedPath.ToString());
                foldernames.Add(Path.GetFileName(dialog.SelectedPath.ToString()));
               
            }
            else
            {
                //assume it is in @"Y:\Product Development\Style Specifications" but not testfolder
                //use linq to find folders matching that description
                testString = findFolder(@"Y:\Product Development\Style Specifications", testString);
                //testString = findFolder(@"Y:\Product Development\Style Specifications\800-899\801", testString);

                // if this string cannot yield a styleID then return user and delete style ID
                if (testString == "")
                { insertStyleId.Text = ""; }

                //add the folder to the list
                if (Directory.Exists(testString))
                {
                    styleFolders.Add(testString);
                    foldernames.Add(Path.GetFileName(testString));
                }
            }

            
            //display to the list for the user
            displayList.ItemsSource = foldernames;
            displayList.Items.Refresh();

            if (displayList.Items.Count != 0)
            {
                if (GoButton.IsVisible == false)
                { GoButton.Visibility = Visibility.Visible; }
            }
        }

        //pass the entire spec list and the folder for all specs
        private void MasterClick(object sender, RoutedEventArgs e)
        {
            //possibly add a warning that this takes like FOREVER

            List<string> specs = new List<string>();
            specs.Add("StyleID");
            specs.Add("Poly");
            specs.Add("Poly Quotes");
            specs.Add("Cardboard");
            specs.Add("Frame");
            specs.Add("Pattern");
            specs.Add("Layouts");
            specs.Add("Sewing");
            specs.Add("Upholstery");
            specs.Add("ProductInfo");
            specs.Add("Dimensions");
            specs.Add("Weights");
            specs.Add("Cartoning");
            specs.Add("Photos");

            //get a datatable
            DataTable table = App.Start();
            //pass table to the current excel sheet
            AlternateExcel.BuildWorkSheet(table, specs);
        }

        //for each checkbox checked add the spec to a list
        private List<string> getCheckedBoxes()
        {
            List<string> specs = new List<string>();
            specs.Add("StyleID"); //always will want this shown
            if(checkBoxPoly.IsChecked == true)
            { specs.Add("Poly"); }
            if(checkBoxProdInfo.IsChecked == true)
            { specs.Add("Poly Quotes"); }
            if(checkBoxCB.IsChecked == true)
            { specs.Add("Cardboard"); }
            if (checkBoxFrame.IsChecked == true)
            { specs.Add("Frame"); }
            if (checkBoxPattern.IsChecked == true)
            { specs.Add("Pattern"); }
            if (checkBoxLayouts.IsChecked == true)
            { specs.Add("Layouts"); }
            if (checkBoxSewing.IsChecked == true)
            { specs.Add("Sewing"); }
            if (checkBoxUphol.IsChecked == true)
            { specs.Add("Upholstery"); }
            if (checkBoxProdInfo.IsChecked == true)
            { specs.Add("ProductInfo"); }
            if (checkBoxDims.IsChecked == true)
            { specs.Add("Dimensions"); }
            if (checkBoxWeights.IsChecked == true)
            { specs.Add("Weights"); }
            if (checkBoxCartoning.IsChecked == true)
            { specs.Add("Cartoning"); }
            if (checkBoxPhotos.IsChecked == true)
            { specs.Add("Photos"); }

            return specs;
        }

        //after selecting files and specs desired. Go button request a datatable of those specs
        //  and filter through those specs for just the ones click and pass that to the 
        //  excel file builder
        private void GoButton_Click(object sender, RoutedEventArgs e)
        {
            //generate a list of specs we'd like based on clicked boxes
            List<string> specs = getCheckedBoxes();

            //use a list made from chosen folders (probably a public list)
            DataTable table = App.Start(styleFolders);

            AlternateExcel.BuildWorkSheet(table, specs);
        }

        //**TODO fix code where it might grab similar style 
        //example grabing 811-5011-91 instead of 811-5011
        //given a root folder and styleID return the path *if it exists
        private string findFolder(string baseFolder, string styleID)
        {
            string targetPath = "";            

            //filter for dwgs
            IEnumerable<FileInfo>  fileList =
                from file in masterList
                where file.Name.Contains(styleID) 
                where ContainsFolder(file.FullName, "TestFolder") == false
                orderby file.DirectoryName
                select file;

            
            if (fileList.Count() > 0)
            {
                foreach (FileInfo fi in fileList)
                {
                    //just using first file in the list
                    //var filePath = fileList.First().FullName;
                    var filePath = fi.FullName;
                    string[] pathParts = filePath.Split(Path.DirectorySeparatorChar);
                    //find the part of the path that has our style ID
                    for (int i = 0; !pathParts[i].Contains(styleID) && i < pathParts.Count(); i++)
                    { targetPath = targetPath + pathParts[i] + @"\"; }
                    targetPath = targetPath + styleID;

                    if (Directory.Exists(targetPath))
                    { break; }
                    else
                    { targetPath = ""; }
                }
                
            }
            return targetPath;
        }

        //check if the file path is under a specific folder
        //NOT USING
        private static bool IsFileBelowDirectory(string fileInfo, string directoryInfo, string separator)
        {
            var directoryPath = string.Format("{0}{1}"
            , directoryInfo
            , directoryInfo.EndsWith(separator) ? "" : separator);

            return fileInfo.StartsWith(directoryPath, StringComparison.OrdinalIgnoreCase);
        }

        //check if the file path has a referance to a specific folder
        private static Boolean ContainsFolder(string directory, string unwantedDirectory)
        {
            String[] directories = directory.Split(Path.DirectorySeparatorChar);
            return directories.ToLookup(i => i.ToLower()).Contains(unwantedDirectory);
        }
    }
}
