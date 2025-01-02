using ClosedXML.Excel;
using ExcelTools.Abstraction;
using ExcelTools.Exceptions;
using System.Text;

namespace ExcelTools.ExcelToHtmlTableConverter
{
    public class ExcelToHtmlTableConverter: ExcelHandlerBase<ExcelToHtmlTableConverterOptions, ExcelToHtmlTableConverterResult>
    {
        public override ExcelToHtmlTableConverterResult Process(ExcelToHtmlTableConverterOptions options)
        {
            Options = options;

            if (!options.Validate())
            {
                throw new ExcelToolsException("Wrong options");
            }

            try
            {
                Convert(Options.FilePath, 1);

                return Result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        public void Convert(string filePath, int sheetNumber)
        {
            var sb = new StringBuilder();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(sheetNumber);
                sb.AppendLine("<table style='border-collapse: collapse;'>");

                foreach (var row in worksheet.RowsUsed())
                {
                    sb.AppendLine("<tr>");

                    foreach (var cell in row.CellsUsed())
                    {
                        if (worksheet.MergedRanges.Any(range => range.Contains(cell) && !range.FirstCell().Address.Equals(cell.Address)))
                        {
                            continue;
                        }

                        var (colspan, rowspan) = GetMergedCellInfo(worksheet, cell.Address.RowNumber, cell.Address.ColumnNumber);

                        sb.AppendLine($"<td style='border: 1px solid black; padding: 5px; font-weight: bold; color: {cell.Style.Font.FontColor}; background-color: {cell.Style.Fill.BackgroundColor};' colspan='{colspan}' rowspan='{rowspan}'>");
                        sb.Append(cell.GetFormattedString());
                        sb.AppendLine("</td>");
                    }

                    sb.AppendLine("</tr>");
                }

                sb.AppendLine("</table>");
            }

            File.WriteAllText(Options.ResultFilePath, sb.ToString());
        }

        /// <summary>
        /// Получение данных о объединенных ячейках 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public (int colspan, int rowspan) GetMergedCellInfo(IXLWorksheet worksheet, int rowIndex, int columnIndex)
        {
            var colspan = 1;
            var rowspan = 1;

            foreach (var range in worksheet.MergedRanges)
            {
                if (!range.Contains(worksheet.Cell(rowIndex, columnIndex))) continue;

                colspan = range.ColumnCount();
                rowspan = range.RowCount();

                break;
            }

            return (colspan, rowspan);
        }

    }
}
