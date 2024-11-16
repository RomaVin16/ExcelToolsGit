using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Text;

namespace ExcelTools.Comparison
{
    public static class ComparisonHelper
    {
        /// <summary>
        /// Получение Id
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="intRowNumber"></param>
        /// <param name="idStrings"></param>
        /// <returns></returns>
        public static string GetId(IXLWorksheet worksheet, int intRowNumber, string[] idStrings)
        {
            var result = new StringBuilder();

            foreach (var columnName in idStrings)
            {
                var cellValue = worksheet.Cell(intRowNumber, columnName).GetValue<string>();

                    result.Append(cellValue);
            }

            return result.ToString();
        }

        /// <summary>
        /// Добавление ID в словарь 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="dictionary"></param>
        /// <param name="idStrings"></param>
        /// <param name="headers"></param>
        public static void AddIds(IXLWorksheet worksheet, Dictionary<string, int> dictionary, string[] idStrings, int[] headers)
        {
            for (var i = worksheet.FirstRowUsed().RowNumber(); i <= worksheet.LastRowUsed().RowNumber(); i++)
            {

                if (idStrings != null)
                {
                    if ((headers != null && headers.Contains(i)) || !CheckId(worksheet, idStrings, i))
                    {
                        continue;
                    }

                    dictionary.Add(GetId(worksheet, i, idStrings), i);
                }
                else
                {
                    dictionary.Add(i.ToString(), i);
                }
            }
        }

        /// <summary>
        /// Проверка ID
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="idStrings"></param>
        /// <param name="rowNumber"></param>
        /// <returns></returns>
        public static bool CheckId(IXLWorksheet worksheet, string[] idStrings, int rowNumber)
        {
            return idStrings.Select(columnName => worksheet.Cell(rowNumber, columnName).GetValue<string>()).All(cellValue => cellValue != "");
        }

        /// <summary>
        /// Добавление комментариев в ячейку
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        public static void InsertCommentInCell(IXLCell cell, string oldValue, string newValue)
        {
            var comment = cell.CreateComment();

            if (decimal.TryParse(oldValue, out var oldDecimalValue) && decimal.TryParse(newValue, out var newDecimalValue))
            {
                var difference = newDecimalValue - oldDecimalValue;

                var sign = difference >= 0 ? "+ " : "- ";

                var result = $"{sign}{Math.Abs(difference)}";

                comment.AddText(result);
            }
else
            {
                comment.AddText($"Исходное значение: {oldValue}, \nНовое значение: {newValue}");
            }
        }

/// <summary>
/// Добавление заголовков в листы 
/// </summary>
/// <param name="sourceWorksheet"></param>
/// <param name="newWorkbook"></param>
/// <param name="headers"></param>
public static void InsertHeadersInList(IXLWorksheet sourceWorksheet, IXLWorkbook newWorkbook, int[] headers)
{
    for (var i = 1; i <= headers.Length; i++) 
    {
        var headerRow = headers[i - 1];

        for (var j = sourceWorksheet.Row(headerRow).FirstCellUsed().Address.ColumnNumber; j <= sourceWorksheet.LastColumnUsed().ColumnNumber(); j++)
        {
            newWorkbook.Worksheet("Changed").Cell(i, j).Value = sourceWorksheet.Cell(headerRow, j).Value;
            newWorkbook.Worksheet("Deleted").Cell(i, j).Value = sourceWorksheet.Cell(headerRow, j).Value;
            newWorkbook.Worksheet("Added").Cell(i, j).Value = sourceWorksheet.Cell(headerRow, j).Value;
                }
    }
        }
    }
}
