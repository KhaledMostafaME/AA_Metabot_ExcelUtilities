using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUtils
{
    class Program
    {
        static void Main(string[] args)
        {
            var wb = @"D:/Test/Test.xlsx";
            var ws = "newSheet";
            var a = new ExcelHelpers();
            //a.CreateWorkbook(wb,"Test");
            //a.AddWorksheet(wb,"newSheet");
            //a.RemoveWorksheet(wb, "Test");
            //a.AddRow(wb, ws, 2, 3);
            //a.AddColumn(wb, ws, "B", 1);
            //a.RemoveColumn(wb, ws, "D");
            a.CreatePivot(wb, ws, "A1:C7", "Pivot","PivotTbl1", new string[] { "numberOfOrders" }, new string[] { "Pastry" }, new string[] { "Month" });
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }
    }
}
