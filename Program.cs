using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfCreate
{
    class Program
    {
        static void Main(string[] args)
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
                using (Project1.ExcelHelper helper = new Project1.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(catalog, "Test.xlsx")))
                    {
                        i = 1;
                        while (true) 
                        {
                            try
                            {
                                helper.Get(sheetNum: i, listNeededJpgs, currentSheet);
                            }
                            catch(Exception ex) { Console.WriteLine(ex.Message); break; }
                        }
                        helper.Save();
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
