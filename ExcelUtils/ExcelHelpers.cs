using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace ExcelUtils
{
    public class ExcelHelpers
    {
        public void CreateWorkbook(string workbookName, string sheetName)
        {
            using (var wb = new XLWorkbook())
            {
                wb.Worksheets.Add(sheetName);
                wb.SaveAs(workbookName);
            }
        }

        public void AddWorksheet(string workbookName, string sheetName)
        {
            using (var wb = new XLWorkbook(workbookName))
            {
                wb.Worksheets.Add(sheetName);
                wb.Save();
            }
        }

        public void RemoveWorksheet(string workbookName, string sheetName)
        {
            using (var wb = new XLWorkbook(workbookName))
            {
                wb.TryGetWorksheet(sheetName, out IXLWorksheet ws);
                ws.Delete();
                wb.Save();
            }
        }

        public void AddRow(string workbookName, string sheetName, int row, int rowsToInsert, bool insertBelow = true)
        {
            using (var wb = new XLWorkbook(workbookName))
            {
                wb.TryGetWorksheet(sheetName, out IXLWorksheet ws);

                if (insertBelow)
                    ws.Row(row).InsertRowsBelow(rowsToInsert);
                else
                    ws.Row(row).InsertRowsAbove(rowsToInsert);

                wb.Save();
            }
        }

        public void DeleteRow(string workbookName, string sheetName, int row)
        {
            using (var wb = new XLWorkbook(workbookName))
            {
                wb.TryGetWorksheet(sheetName, out IXLWorksheet ws);
                ws.Row(row).Delete();
                wb.Save();
            }
        }

        public void AddColumn(string workbookName, string sheetName, string col, int ColsToInsert, bool insertAfter = true)
        {
            using (var wb = new XLWorkbook(workbookName))
            {
                wb.TryGetWorksheet(sheetName, out IXLWorksheet ws);

                if (insertAfter)
                    ws.Column(col).InsertColumnsAfter(ColsToInsert);
                else
                    ws.Column(col).InsertColumnsBefore(ColsToInsert);

                wb.Save();
            }
        }

        public void RemoveColumn(string workbookName, string sheetName, string col)
        {
            using (var wb = new XLWorkbook(workbookName))
            {
                wb.TryGetWorksheet(sheetName, out IXLWorksheet ws);

                ws.Column(col).Delete();

                wb.Save();
            }
        }

        public void CreatePivot(string workbookName, string dataSheetName, string dataRange, string pivotSheetName, string pivotName, int pivotStartCol, int pivotStartRow, string[] pivotValues, string[] pivotRows, string[] pivotCols)
        {
            using (var wb = new XLWorkbook(workbookName))
            {
                wb.TryGetWorksheet(dataSheetName, out IXLWorksheet ws);
                var data = ws.Range(dataRange);

                var ptSheet = wb.Worksheets.Add(pivotSheetName);

                var pt = ptSheet.PivotTables.Add(pivotName, ptSheet.Cell(pivotStartRow, pivotStartCol), data);

                foreach (var row in pivotRows)
                    pt.RowLabels.Add(row);

                foreach (var col in pivotCols)
                    pt.ColumnLabels.Add(col);

                foreach (var value in pivotValues)
                    pt.Values.Add(value);

                wb.Save();
            }
        }

        public void SheetAddNamedRange(string workbookName, string sheetName, string range, string rangeName)
        {
            using (var wb = new XLWorkbook(workbookName))
            {
                wb.TryGetWorksheet(sheetName, out IXLWorksheet ws);
                ws.Range(range).AddToNamed(rangeName);
                wb.Save();
            }
        }

        public int GetColNumber(string cell)
        {
            cell = Regex.Replace(cell, @"[\d]", string.Empty);
            cell = cell.ToUpperInvariant();
            int num = 0;
            for (int i = 0; i < cell.Length; i++)
            {
                num *= 26;
                num += cell[i] - 65 + 1;
            }
            return num;
        }

        public void AddConditionalFormatting(string workbookName, string sheetName, string range, string rule)
        {
            using (var wb = new XLWorkbook(workbookName))
            {
                wb.TryGetWorksheet(sheetName, out IXLWorksheet ws);
                //ws.Range(range).AddConditionalFormat().WhenEqualOrGreaterThan(rule);
                ws.Range(range).AddConditionalFormat().WhenEquals(rule);
                wb.Save();
            }
        }
    }
}
