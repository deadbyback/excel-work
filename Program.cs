using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace RKSM_5
{
    class Program
    {
        public static ExcelWork ExcelWork { get; private set; }

        static void Main(string[] args)
        {
            ExcelWork = new ExcelWork();
            Console.WriteLine("\nMake your choice:");
            int choice = Convert.ToInt32(Console.ReadLine());
            switch (choice)
            {
                case 1:
                    Console.WriteLine("\nFill it:");
                    ExcelWork.ReadExcelFile();
                    break;
                case 2:
                    Console.WriteLine("\nFill it:");
                    Console.Write("Class Name:\t");
                    string ClassName = Convert.ToString(Console.ReadLine());
                    Console.Write("Date in format dd.MM.yyyy:\t");
                    DateTime Date = DateTime.Parse(Console.ReadLine());
                    Console.WriteLine("\nNow fill a new data:");
                    Console.Write("Class Name:\t");
                    string NewClassName = Convert.ToString(Console.ReadLine());
                    Console.Write("Student Count:\t");
                    int NewStudentCount = Convert.ToInt32(Console.ReadLine());
                    Console.Write("Date in format dd.MM.yyyy:\t");
                    DateTime NewDate = DateTime.Parse(Console.ReadLine());
                    Console.Write("Group:\t");
                    string NewGroup = Convert.ToString(Console.ReadLine());
                    ExcelWork.UpdateRowCell(ClassName, Date, NewClassName, NewStudentCount, NewDate, NewGroup);
                    break;
                case 3:
                    Console.WriteLine("\nFill it:");
                    Console.Write("Class Name:\t");
                    ClassName = Convert.ToString(Console.ReadLine());
                    Console.Write("Student Count:\t");
                    int StudentCount = Convert.ToInt32(Console.ReadLine());
                    Console.Write("Date:\t");
                    Date = DateTime.Now;
                    Console.Write("Group:\t");
                    string Group = Convert.ToString(Console.ReadLine());
                    ExcelWork.AddNewRowToExcelFile(ClassName, StudentCount, Date, Group);
                    break;
                case 4:
                    Console.WriteLine("\nFill it:");
                    Console.Write("Class Name:\t");
                    ClassName = Convert.ToString(Console.ReadLine());
                    Console.Write("Date in format dd.MM.yyyy:\t");
                    Date = DateTime.Parse(Console.ReadLine());
                    ExcelWork.DeleteRowCell(ClassName, Date);
                    break;
                case 5:
                    Console.WriteLine("\nChoose column for sorting:");
                    Console.WriteLine("\nA - for Class Name\tB - for Student Count\nC - for Date\tD - for Group");
                    Console.Write("Your choice: \t");
                    string Path = Console.ReadLine();
                    Console.Write("Ascending? (Yes or No): \t");
                    string AscResponse = Console.ReadLine();
                    AscResponse.ToLower(); 
                    bool Ascending = true;
                    if (AscResponse == "no")
                    {
                        Ascending = false;
                    }
                    ExcelWork.SortData(Path, Ascending);
                    break;
                default:
                    break;
            }
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
            if (!System.IO.File.Exists(filePath))
            {
                Console.WriteLine("\nFile doesn't find. ");
                Console.WriteLine("\nCreating new Excel file...");
                CreateExcelFile();
            }

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nReading the Excel File...");
            Console.BackgroundColor = ConsoleColor.Black;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Range xlRange = xlWorkSheet.UsedRange;
            int totalRows = xlRange.Rows.Count;
            int totalColumns = xlRange.Columns.Count;

            string firstValue, secondValue, forthValue;
            int thirdValue;

            for (int rowCount = 1; rowCount <= totalRows; rowCount++)
            {

                firstValue = Convert.ToString((xlRange.Cells[rowCount, 1] as Excel.Range).Text);
                secondValue = Convert.ToInt32((xlRange.Cells[rowCount, 2] as Excel.Range).Text);
                thirdValue = Convert.ToDateTime((xlRange.Cells[rowCount, 3] as Excel.Range).Text);
                forthValue = Convert.ToString((xlRange.Cells[rowCount, 4] as Excel.Range).Text);

                Console.WriteLine(firstValue + "\t" + secondValue + "\t" + thirdValue + "\t" + forthValue);

            }

            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("End of the file...");
        }

        public static void AddNewRowToExcelFile(string ClassName, int StudentCount, DateTime Date, string Group)
        {
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
                Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
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
            Console.WriteLine("\nRecord added successfully...");
            Console.BackgroundColor = ConsoleColor.Black;
        }

        public static void UpdateRowCell(string oldClassName, DateTime oldDate, string newClassName, int newStudentCount, DateTime newDate, string newGroup)
        {
            if (!System.IO.File.Exists(filePath))
            {
                Console.WriteLine("\nFile doesn't find. ");
                Console.WriteLine("\nCreating new Excel file...");
                CreateExcelFile();
            }

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nReading the Excel File...");
            Console.BackgroundColor = ConsoleColor.Black;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Range xlRange = xlWorkSheet.UsedRange;
            int totalRows = xlRange.Rows.Count;

            for (int i = 1; i <= totalRows; i++)
            {
                Excel.Range ClassName = xlRange.Cells[i, 1] as Excel.Range;
                Excel.Range Date = xlRange.Cells[i, 3] as Excel.Range;

                if ((string)ClassName.Value2 == oldClassName && (DateTime)Date == oldDate)
                {
                    xlRange.Cells[i, 1].Value = newClassName;
                    xlRange.Cells[i, 2].Value = newStudentCount;
                    xlRange.Cells[i, 3].Value = newDate;
                    xlRange.Cells[i, 4].Value = newGroup;
                }
            }

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nExcel File has been updated!");
            Console.BackgroundColor = ConsoleColor.Black;
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

        public static void DeleteRowCell(string ClassName, DateTime Date)
        {
            if (!System.IO.File.Exists(filePath))
            {
                Console.WriteLine("\nFile doesn't find. ");
                Console.WriteLine("\nCreating new Excel file...");
                CreateExcelFile();
            }

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nReading the Excel File...");
            Console.BackgroundColor = ConsoleColor.Black;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Range xlRange = xlWorkSheet.UsedRange;
            int totalRows = xlRange.Rows.Count;

            for (int i = 1; i <= totalRows; i++)
            {
                Excel.Range ClassNameCell = xlRange.Cells[i, 1] as Excel.Range;
                Excel.Range DateCell = xlRange.Cells[i, 3] as Excel.Range;

                if ((string)ClassNameCell.Value2 == ClassName && (DateTime)DateCell == Date)
                {
                    Excel.Range range = xlWorkSheet.get_Range("A"+i, "D"+i);
                    range.EntireRow.Delete(Type.Missing);
                }
            }
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nExcel File has been updated!");
            Console.BackgroundColor = ConsoleColor.Black;
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


        public static void SortData(string Path, bool Ascending = true)
        {
            if (!System.IO.File.Exists(filePath))
            {
                Console.WriteLine("\nFile doesn't find. ");
                Console.WriteLine("\nCreating new Excel file...");
                CreateExcelFile();
            }

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nReading the Excel File...");
            Console.BackgroundColor = ConsoleColor.Black;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.UsedRange.Select();
            xlWorkSheet.Sort.SortFields.Clear();
            if (Ascending == true)
            {
                xlWorkSheet.Sort.SortFields.Add(xlWorkSheet.UsedRange.Columns[Path], Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending, System.Type.Missing, Excel.XlSortDataOption.xlSortNormal);
            }
            else
            {
                xlWorkSheet.Sort.SortFields.Add(xlWorkSheet.UsedRange.Columns[Path], Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending, System.Type.Missing, Excel.XlSortDataOption.xlSortNormal);
            }

            var sort = xlWorkSheet.Sort;
            sort.SetRange(xlWorkSheet.UsedRange);
            sort.Header = Excel.XlYesNoGuess.xlYes;
            sort.MatchCase = false;
            sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
            sort.SortMethod = Excel.XlSortMethod.xlPinYin;
            sort.Apply();

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nExcel File has been sorted!");
            Console.BackgroundColor = ConsoleColor.Black;
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
