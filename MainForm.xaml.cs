using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace PdfCreate
{
    /// <summary>
    /// Логика взаимодействия для Main.xaml
    /// </summary>
    public partial class MainForm : System.Windows.Controls.UserControl
    {
        private string _excelFilePath;
        private string _folderPath;
        private string _saveFolder;

        public MainForm()
        {
            InitializeComponent();
        }
        
        private void Folder_Dialog(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            /*dialog.RootFolder = Environment.CurrentDirectory; // Use current value for initial dir
            dialog. = "Выберите папку с изображениями"; // instead of default "Save As"
            dialog.Filter = "Directory|*.this.directory"; // Prevents displaying files
            dialog.FileName = "select"; // Filename will then be "select.this.directory"*/

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string path = dialog.SelectedPath;
                /*// Remove fake filename from resulting path
                path = path.Replace("\\select.this.directory", "");
                path = path.Replace(".this.directory", "");
                // If user has changed the filename, create the new directory*/
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                // Our final value is in path
                _folderPath = path;

                CatalogName.Text = _folderPath;
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

        void DrawImage(XGraphics gfx, string jpegSamplePath, int x, int y, int width, int height)
        {
            XImage image = XImage.FromFile(jpegSamplePath);
            gfx.DrawImage(image, x, y, width, height);
        }

        private void ProgramLogic(object sender, RoutedEventArgs e)
        {
            string catalog = _folderPath, file = _excelFilePath, currentSheet = null;
            int n = 0;
            List<string> listNeededJpgs = new List<string>();
            List<List<string>> listOfPJpgs = new List<List<string>>();
            var fileNames = new List<string> { };

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
                    int ab = names.Length - 2;
                    listOfPJpgs[n][0] = names[ab];
                    listOfPJpgs[n][1] = FI.FullName;
                    n++;
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
            }

            try
            {
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: file))
                    {
                        int i = 1;
                        int flag = 2;
                        while (true)
                        {
                            var document = new PdfDocument();

                            try
                            {
                                listNeededJpgs.Clear();
                                fileNames.Clear();
                                if (!helper.Get(sheetNum: i, listNeededJpgs, ref currentSheet)) break;

                                foreach (string need in listNeededJpgs)
                                {
                                    for (int j = 0; j < n; j++)
                                    {
                                        if (need == listOfPJpgs[j][0])
                                        {
                                            fileNames.Add(listOfPJpgs[j][1]);
                                            break;
                                        }
                                        if (j == n - 1 && flag ==2)
                                        {
                                            var result = System.Windows.Forms.MessageBox.Show("Не все сканы доступны для объединения. Вы хотите продолжть в любом случае? ", "Недостаточно файлов",
                                                                    MessageBoxButtons.YesNo,
                                                                    MessageBoxIcon.Question);
                                            if (result == DialogResult.No)
                                            {
                                                flag = 0;
                                                break;
                                            } else flag = 1;
                                        }
                                    }
                                    if (flag == 0) break;   
                                }
                            }
                            catch (Exception ex) { Console.WriteLine(ex.Message); break; }
                            if (flag != 0)
                            {
                                foreach (string fileName in fileNames)
                                {
                                    var page = document.AddPage();
                                    XGraphics gfx = XGraphics.FromPdfPage(page);
                                    DrawImage(gfx, fileName, 0, 0, (int)page.Width, (int)page.Height);
                                }
                            }
                            else break;

                            if (document.PageCount > 0) document.Save(_saveFolder + "\\" + currentSheet + ".pdf");
                            i++;

                        }
                        helper.Save();
                        var result1 = System.Windows.Forms.MessageBox.Show("Файлы pdf вы можете найти в выбранной папке.", "Готово",
                                                                    MessageBoxButtons.OK,
                                                                    MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        private void Save_Folder(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            /*dialog.RootFolder = Environment.CurrentDirectory; // Use current value for initial dir
            dialog. = "Выберите папку с изображениями"; // instead of default "Save As"
            dialog.Filter = "Directory|*.this.directory"; // Prevents displaying files
            dialog.FileName = "select"; // Filename will then be "select.this.directory"*/

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string path = dialog.SelectedPath;
                /*// Remove fake filename from resulting path
                path = path.Replace("\\select.this.directory", "");
                path = path.Replace(".this.directory", "");
                // If user has changed the filename, create the new directory*/
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                // Our final value is in path
                _saveFolder = path;

                SaveFolder.Text = _saveFolder;
            }
        }
    }
}
