using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using Mmu.Mlh.ExcelAccess.Areas.Models;
using OfficeOpenXml;

using models = Mmu.Mlh.ExcelAccess.Areas.Models;

namespace Mmu.Mlh.ExcelAccess.Areas.Services.Implementation
{
    internal class ExcelWorkbookFactory : IExcelWorkbookFactory
    {
        private readonly IFileSystem _fileSystem;

        public ExcelWorkbookFactory(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        public models.ExcelWorkbook Create(Stream stream)
        {
            var excelPackage = new ExcelPackage(stream);
            var worksheets = new List<models.ExcelWorksheet>();

            foreach (var ws in excelPackage.Workbook.Worksheets)
            {
                var cells = ws.Cells.Where(cell => cell.Value != null).Select(cell => AdaptCell(cell)).ToList();
                worksheets.Add(new models.ExcelWorksheet(ws.Name, cells));
            }

            var result = new models.ExcelWorkbook(worksheets);
            return result;
        }

        public models.ExcelWorkbook Create(string filePath)
        {
            using (var stream = _fileSystem.File.OpenRead(filePath))
            {
                return Create(stream);
            }
        }

        private static ExcelCell AdaptCell(ExcelRangeBase cell)
        {
            var adr = new models.ExcelCellAddress(cell.Start.Row, cell.End.Column);
            var rgb = cell.Style.Fill.BackgroundColor.Rgb;

            var color = Color.Transparent;
            if (!string.IsNullOrEmpty(rgb))
            {
                var argb = int.Parse(rgb, NumberStyles.HexNumber);
                color = Color.FromArgb(argb);
            }

            var style = new ExcelCellStyle(color);
            var excelCell = new ExcelCell(cell.Value, adr, style);
            return excelCell;
        }
    }
}