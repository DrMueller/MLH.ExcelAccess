using System.IO;
using OfficeOpenXml;
using models = Mmu.Mlh.ExcelAccess.Areas.Models;

namespace Mmu.Mlh.ExcelAccess.Areas.Services.Implementation
{
    internal class ExcelFileWriter : IExcelFileWriter
    {
        public void WriteToStream(models.ExcelWorkbook workbook)
        {
            var excelPackage = new ExcelPackage();

            foreach (var worksheet in workbook.Worksheets)
            {
                var ws = excelPackage.Workbook.Worksheets.Add(worksheet.Name);
                foreach (var cell in worksheet.Cells)
                {
                    var nativeCell = ws.Cells[cell.Address.Row, cell.Address.Column];
                    nativeCell.Value = cell.Value;
                    nativeCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    nativeCell.Style.Fill.BackgroundColor.SetColor(cell.Style.BackgroundColor);
                }
            }

            var fi = new FileInfo(@"C:\Users\mlm\Desktop\Test.xlsx");
            excelPackage.SaveAs(fi);
        }
    }
}