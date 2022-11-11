using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Aspose.Words;

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

        void ProgramLogic(string[] args)
        {
            string catalog = _folderPath, file = _excelFilePath, currentSheet = null;
            int n = 0;
            List<string> listNeededJpgs = new List<string>();
            List<List<string>> listOfPJpgs = new List<List<string>>();

            foreach (string findedFile in Directory.EnumerateFiles(catalog, "*.jpg*", SearchOption.AllDirectories))
            {
                FileInfo FI;
                try
                {
                    FI = new FileInfo(findedFile);
                    listOfPJpgs.Add(new List<string>());
                    listOfPJpgs[n].Add(" ");
                    listOfPJpgs[n].Add(" ");
                    string[] names = FI.Name.Split(' ', '.');
                    listOfPJpgs[n][0] = names[1];
                    listOfPJpgs[n][1] = FI.FullName;
                    n++;
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
            }

            try
            {
                var doc = new Document();
                var builder = new DocumentBuilder(doc);
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: file))
                    {
                        int i = 1;
                        while (true)
                        {
                            try
                            {
                                helper.Get(sheetNum: i, listNeededJpgs, currentSheet);
                                var fileNames = new List<string> { };

                                foreach (string need in listNeededJpgs)
                                {
                                    for (int j = 0; j <= n; j++)
                                    {
                                        if(need == listOfPJpgs[j][0])
                                        {
                                            fileNames.Add(listOfPJpgs[j][1]);
                                            break;
                                        }
                                    }
                                }

                                foreach (string fileName in fileNames)
                                {
                                    builder.InsertImage(fileName);
                                    builder.Writeln();
                                }

                                doc.Save(currentSheet + ".pdf");

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
