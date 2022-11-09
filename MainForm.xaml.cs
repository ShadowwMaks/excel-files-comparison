using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace PdfCreate
{
    /// <summary>
    /// Логика взаимодействия для Main.xaml
    /// </summary>
    public partial class MainForm : UserControl
    {
        private string _excelFilePath;
        private string _folderPath;

        public MainForm()
        {
            InitializeComponent();
        }
        
        private void Folder_Dialog(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.InitialDirectory = Environment.CurrentDirectory; // Use current value for initial dir
            dialog.Title = "Выберите папку с изображениями"; // instead of default "Save As"
            dialog.Filter = "Directory|*.this.directory"; // Prevents displaying files
            dialog.FileName = "select"; // Filename will then be "select.this.directory"

            if (dialog.ShowDialog() == true)
            {
                string path = dialog.FileName;
                // Remove fake filename from resulting path
                path = path.Replace("\\select.this.directory", "");
                path = path.Replace(".this.directory", "");
                // If user has changed the filename, create the new directory
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                // Our final value is in path
                _folderPath = path;

                FileName.Text = _folderPath;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                _excelFilePath = filename;

                FileName.Text = _excelFilePath;
            }
        }

        static void ProgramLogic(string[] args)
        {
            string catalog = Environment.CurrentDirectory, currentSheet = null;
            int i = 0;
            List<string> listNeededJpgs = new List<string>();
            List<List<string>> listOfPJpgs = new List<List<string>>();

            foreach (string findedFile in Directory.EnumerateFiles(catalog, "*.jpg*", SearchOption.AllDirectories))
            {
                FileInfo FI;
                try
                {
                    FI = new FileInfo(findedFile);
                    listOfPJpgs.Add(new List<string>());
                    listOfPJpgs[i][0] = FI.Name;
                    listOfPJpgs[i][1] = FI.FullName;
                    i++;
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
            }

            try
            {
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: System.IO.Path.Combine(catalog, "Test.xlsx")))
                    {
                        i = 1;
                        while (true)
                        {
                            try
                            {
                                helper.Get(sheetNum: i, listNeededJpgs, currentSheet);
                            }
                            catch (Exception ex) { Console.WriteLine(ex.Message); break; }
                        }
                        helper.Save();
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
