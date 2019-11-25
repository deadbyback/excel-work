using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
/*
 using Excel = Microsoft.Office.Interop.Excel;
     */

namespace RKSM_5
{
    class Program
    {
        public static ExcelWork ExcelWork { get; private set; }

        static void Main(string[] args)
        {
            ExcelWork = new ExcelWork();
            ExcelWork.CreateExcelFile();
            ExcelWork.ReadExcelFile();
            ExcelWork.AddNewRowsToExcelFile();
            ExcelWork.ReadExcelFile();
            ExcelWork.DeleteRowCellFromExcelFile();
            ExcelWork.ReadExcelFile();
            Console.Read();
        }
    }


    class ExcelWork
    {
        public static string filePath = @"I:\RKSM_5_1\data.xlsx";

        public static void CreateExcelFile()
        {
            Excel.Application xlApp = new Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not installed in the system...");
                return;
            }

            object misValue = System.Reflection.Missing.Value;

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Class Name";
            xlWorkSheet.Cells[1, 2] = "Count";
            xlWorkSheet.Cells[1, 3] = "Date";
            xlWorkSheet.Cells[1, 4] = "Group";
            xlWorkSheet.Cells[2, 1] = "IT-Archicture";
            xlWorkSheet.Cells[2, 2] = 9;
            xlWorkSheet.Cells[2, 3] = DateTime.Now;
            xlWorkSheet.Cells[2, 4] = "CS-19sm";

            xlWorkBook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("Excel file created successfully...");
            Console.BackgroundColor = ConsoleColor.Black;
        }

        public static void ReadExcelFile()
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nReading the Excel File...");
            Console.BackgroundColor = ConsoleColor.Black;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Range xlRange = xlWorkSheet.UsedRange;
            int totalRows = xlRange.Rows.Count;
            int totalColumns = xlRange.Columns.Count;

            string firstValue, secondValue, thirdValue, forthValue;

            for (int rowCount = 1; rowCount <= totalRows; rowCount++)
            {

                firstValue = Convert.ToString((xlRange.Cells[rowCount, 1] as Excel.Range).Text);
                secondValue = Convert.ToString((xlRange.Cells[rowCount, 2] as Excel.Range).Text);
                thirdValue = Convert.ToString((xlRange.Cells[rowCount, 2] as Excel.Range).Text);
                forthValue = Convert.ToString((xlRange.Cells[rowCount, 2] as Excel.Range).Text);

                Console.WriteLine(firstValue + "\t" + secondValue + "\t" + thirdValue + "\t" + forthValue);

            }

            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("End of the file...");
        }

        public static void AddNewRowsToExcelFile(string ClassName, int StudentCount, DateTime Date, string Group)
        {
            //        IList<Attend> List = new List<Attend>() {
            //    new Attend(){ ID=1003, Name="Indraneel"},
            //    new Attend(){ ID=1004, Name="Neelohith"},
            //    new Attend(){ ID=1005, Name="Virat"}
            //};
            IList<Attend> List = new List<Attend>()
            {
                new Attend(){
                    ClassName = ClassName,
                    StudentCount = StudentCount,
                    Date = Date,
                    Group = Group
                }
            };

            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Range xlRange = xlWorkSheet.UsedRange;
            int rowNumber = xlRange.Rows.Count + 1;

            foreach (Attend les in List)
            {
                xlWorkSheet.Cells[rowNumber, 1] = les.ClassName;
                xlWorkSheet.Cells[rowNumber, 2] = les.StudentCount;
                xlWorkSheet.Cells[rowNumber, 3] = les.Date;
                xlWorkSheet.Cells[rowNumber, 4] = les.Group;
                rowNumber++;
            }

            // Disable file override confirmaton message  
            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);
            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nRecords Added successfully...");
            Console.BackgroundColor = ConsoleColor.Black;
        }

        public static void DeleteRowCellFromExcelFile()
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nDeleting the Records...");
            Console.BackgroundColor = ConsoleColor.Black;

            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Range range1 = xlWorkSheet.get_Range("A2", "B2");

            // To Delete Entire Row - below rows will shift up  
            range1.EntireRow.Delete(Type.Missing);

            Excel.Range range2 = xlWorkSheet.get_Range("B3", "B3");
            range2.Cells.Clear();

            // To Delete Cells - Below cells will shift up  
            // range2.Cells.Delete(Type.Missing);  

            // Disable file override confirmaton message  
            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);
            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }


    public class Attend
    {
        public string ClassName { get; set; }
        public int StudentCount { get; set; }
        public DateTime Date { get; set; }
        public string Group { get; set; }
    }
}
