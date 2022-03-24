using System;
using Microsoft.Office.Interop.Excel;

class Program
{
    public static void Main(string[] args)
    {
        Application excelApp = new Application();


        if (excelApp == null)
        {
            Console.WriteLine("Excel не обнаружен");
            return;
        }

        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\Izagakhmaevra\Desktop\Excel\TestExel.xlsx");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;

        for (int i = 1; i <= rows; i++)
        {
            Console.Write("\r\n");
            for (int j = 1; j <= cols; j++)
            {

                if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");
            }
        }
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        Console.ReadLine();
    }
}