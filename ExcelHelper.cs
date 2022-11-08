using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Project1
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

        internal bool Get(int sheetNum, List<string> listNeededJpgs, string currentSheet)
        {
            Excel.Worksheet ObjWorkSheet;
            try
            {
                ObjWorkSheet = (Excel.Worksheet)_workbook.Sheets[sheetNum];
                currentSheet = ObjWorkSheet.Name;
                /*как-то надо выбрать правильный столбец и получить номера*/
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
    }
}
