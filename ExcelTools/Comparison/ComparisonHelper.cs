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

        public void AddIds(IXLWorksheet worksheet, HashSet<string> hash, string[] idStrings)
        {
            for (var i = worksheet.FirstRowUsed().RowNumber(); i <= worksheet.LastRowUsed().RowNumber(); i++)
            {
                hash.Add(GetId(worksheet, i, idStrings));
            }
        }
    }
}
