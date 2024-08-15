using ClosedXML.Excel;

namespace ExcelTools
{
    public class Cleaner
    {
        /// <summary>
        /// Удаление только пустых строк в конкретном листе
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        public void DeleteEmptyRowsInSheet(IXLWorksheet item, CleanResult result)
        {
            ExcelHelper.CheckingTheFileForEmptiness(item, result);

            for (var i = item.LastRowUsed().RowNumber(); i >= 1; i--)
            {
                result.RowsProcessed++;
                if (item.Row(i).IsEmpty())
                {
                    item.Row(i).Delete();
                    result.RowsRemoved++;
                }
            }
        }

        /// <summary>
        /// Удаление пустых строк, а также содержащих табуляции, пробелы и другие невидимые символы
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        public void DeleteNonPrintableOrEmptyRowsInSheet(IXLWorksheet item, CleanResult result)
        {
            for (var i = item.LastRowUsed().RowNumber(); i >= 1; i--)
            {
                result.RowsProcessed++;

                if (item.Row(i).IsEmpty())
                {
                    item.Row(i).Delete();
                    result.RowsRemoved++;
                }
                else
                {
                    foreach (var cell in item.Row(i).Cells())
                    {
                        if (string.IsNullOrWhiteSpace(cell.Value.ToString()) || item.Row(i).IsEmpty())
                        {
                            item.Row(i).Delete();
                            result.RowsRemoved++;
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// удаление пустых строк в Excel файле
        /// </summary>
        public CleanResult DeleteEmptyStringInWorkbook(CleanOptions cleanerOptions)
        {
            var cleanResult = new CleanResult();

            var workbook = ExcelHelper.WriteFile(cleanerOptions.FilePath + ".xlsx", cleanResult);

            foreach (var item in workbook.Worksheets)
            {
                if (cleanerOptions.CleanWhiteSpaceRows)
                {
                    DeleteNonPrintableOrEmptyRowsInSheet(item, cleanResult);
                }
                else
                {
                    DeleteEmptyRowsInSheet(item, cleanResult);
                }
            }

            workbook.SaveAs(cleanerOptions.ResultFilePath + ".xlsx");

            return cleanResult;
        }
    }
}
