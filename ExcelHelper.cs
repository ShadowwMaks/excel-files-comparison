using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace PdfCreate
{
    internal class ExcelHelper: IDisposable
    {
        private Application _excel;
        private Workbook _workbook;
        private string _filePath;

        public ExcelHelper() 
        {
            _excel = new Excel.Application();
        }

        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                }
                else
                {
                    Console.WriteLine("Файла не существует");
                }
                return true;
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal bool Set(string column, int row, object data)
        {
            try
            {
                (_excel.ActiveSheet as Excel.Worksheet).Cells[row, column] = data;
                return true;
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }
            return false; 
        }

        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
            }
            else
            {
                _workbook.Save();
            }
        }

        internal bool Get(int sheetNum, List<string> listNeededJpgs, ref string nameSheet)
        {
            Excel.Worksheet ObjWorkSheet;
            try
            {
                ObjWorkSheet = (Excel.Worksheet)_workbook.Sheets[sheetNum];
                nameSheet = ObjWorkSheet.Name;
                object data;
                string nakl;
                int i = 11;
                while (true)
                {
                    char[] separators = new char[] { ' ', '.' , 'Д' };
                    data = ObjWorkSheet.Cells[i, "F"].Formula2Local;
                    nakl = data.ToString();
                    if (!string.IsNullOrEmpty(nakl))
                    {
                        string[] num = nakl.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        listNeededJpgs.Add(num[0]);
                        i++;
                    }
                    else
                    {
                        break;
                    }
                }
                return true;
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }
        public void Dispose()
        {
            try
            {
                _workbook.Close();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        public void GetWholeColumn()
        {
            var xlWorksheet = (Excel.Worksheet)_workbook.Worksheets.get_Item(1);

            object[] columnValue = xlWorksheet.Range["C"].Value2;
        }
    }
}
