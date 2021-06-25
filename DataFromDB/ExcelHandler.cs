using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace DataFromDB
{
    class ExcelHandler
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = Missing.Value;
        int row = 2;


        internal void CreatExcelObject()
        {
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "MFN";
            xlWorkSheet.Cells[1, 2] = "Автор";
            xlWorkSheet.Cells[1, 3] = "Заглавие";
            xlWorkSheet.Cells[1, 4] = "Место";
            xlWorkSheet.Cells[1, 5] = "Год";
            xlWorkSheet.Cells[1, 6] = "Кол-во экземпляров";
            xlWorkSheet.Cells[1, 7] = "Первый инв. номер";

            xlWorkSheet.get_Range("A1", "G1").Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

        }

        internal void AddRow(BriefDiscription brief)
        {
            xlWorkSheet.Cells[row, 1] = brief.Mfn;
            xlWorkSheet.Cells[row, 2] = brief.Autor;
            xlWorkSheet.Cells[row, 3] = brief.Title;
            xlWorkSheet.Cells[row, 4] = brief.Location;
            xlWorkSheet.Cells[row, 5] = brief.Year;
            xlWorkSheet.Cells[row, 6] = brief.NumberOfCopies;
            xlWorkSheet.Cells[row, 7] = brief.FirstInvNum;
            row++;
        }
        
        

        internal void SaveFile() {
            string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            xlWorkBook.SaveAs(appPath + @"\ListOfRecords.xls",
                Excel.XlFileFormat.xlWorkbookNormal,
                misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive,
                misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
