using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace ExcelTools.FileInfo
{
    public class FileInfoResult
    {
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }

        public FileInfoResult(string filePath, int worksheetNumber)
        {
            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = document.WorkbookPart;
                var sheets = workbookPart?.Workbook.Sheets?.Elements<Sheet>().ToList();

                if (worksheetNumber < 1 || worksheetNumber > sheets.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(worksheetNumber), "Worksheet number is out of range.");
                }

                var sheet = sheets[worksheetNumber - 1];
                var worksheetPart = (WorksheetPart)workbookPart?.GetPartById(sheet.Id);

                var rowsCount = 0;
                var columnsCount = 0;

                using (var reader = OpenXmlReader.Create(worksheetPart))
                {
                    while (reader.Read())
                    {
                        if (reader.ElementType != typeof(Row)) continue;

                        rowsCount++;

                        var row = (Row)reader.LoadCurrentElement()!;

                        columnsCount = Math.Max(columnsCount, row.Elements<Cell>().Count());
                    }
                }

                RowCount = rowsCount;
                ColumnCount = columnsCount;
            }
        }

    }
}
