using ClosedXML.Excel;

namespace ExcelTools.Helpers
{
    public static class ExcelHelper
    {
        /// <summary>
        /// Удаление пустых столбцов 
        /// </summary>
        /// <param name="worksheet"></param>
        public static void DeleteEmptyColumns(IXLWorksheet worksheet)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

            for (var i = lastColumn; i >= 1; i--)
            {
                if (worksheet.Column(i).IsEmpty())
                {
                    worksheet.Column(i).Delete();
                }
            }
        }
    }
}