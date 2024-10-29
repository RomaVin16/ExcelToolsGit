using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTools.Comparison
{
    public class ComparisonHelper
    {
        /// <summary>
        /// Получение Id
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="intRowNumber"></param>
        /// <param name="idStrings"></param>
        /// <returns></returns>
        public string GetId(IXLWorksheet worksheet, int intRowNumber, string[] idStrings)
        {
            var result = new StringBuilder();

            foreach (var columnName in idStrings)
            {
                var cellValue = worksheet.Cell(intRowNumber, columnName).GetValue<string>();

                    result.Append(cellValue);
            }

            return result.ToString();
        }

        public void AddIds(IXLWorksheet worksheet, HashSet<string> hash, string[] idStrings, int[] headers)
        {
            for (var i = worksheet.FirstRowUsed().RowNumber(); i <= worksheet.LastRowUsed().RowNumber(); i++)
            {
                if ((headers != null && headers.Contains(i)) || !CheckId(worksheet, idStrings, i))
                {
                    continue;
                }

                hash.Add(GetId(worksheet, i, idStrings));
            }
        }

        public bool CheckId(IXLWorksheet worksheet, string[] idStrings, int rowNumber)
        {
            return idStrings.Select(columnName => worksheet.Cell(rowNumber, columnName).GetValue<string>()).All(cellValue => cellValue != "");
        }

        public void InsertCommentInCell(IXLCell cell, string oldValue, string newValue)
        {
            var comment = cell.CreateComment();
            comment.AddText($"Исходное значение: {oldValue}, Новое значение: {newValue}");
        }
    }
}
