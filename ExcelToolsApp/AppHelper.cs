using ExcelTools.Abstraction;
using ExcelTools.Cleaner;
using ExcelTools.ColumnSplitter;
using ExcelTools.DuplicateRemover;
using ExcelTools.Merger;
using ExcelTools.Splitter;

namespace ExcelTools.App
{
    public static class AppHelper
    {
        /// <summary>
        /// Вывод статистики работы программы на консоль
        /// </summary>
        /// <param name="result"></param>
        public static void PrintProgramOperationStatistics(CleanResult result)
        {
            if (result.Code == ResultCode.Error)
            {
                Console.WriteLine("Error occured: " + result.ErrorMessage);
            }
            else
            {
                Console.WriteLine("Rows processed: " + result.RowsProcessed);
                Console.WriteLine("Rows removed: " + result.RowsRemoved);
            }
        }

        /// <summary>
        /// Вывод статистики работы программы на консоль
        /// </summary>
        /// <param name="result"></param>
        public static void PrintProgramOperationStatistics(DuplicateRemoverResult result)
        {
            if (result.Code == ResultCode.Error)
            {
                Console.WriteLine("Error occured: " + result.ErrorMessage);
            }
            else
            {
                Console.WriteLine("Rows processed: " + result.RowsProcessed);
                Console.WriteLine("Rows removed: " + result.RowsRemoved);
            }
        }

        /// <summary>
        /// Вывод статистики работы программы на консоль
        /// </summary>
        /// <param name="result"></param>
        public static void PrintProgramOperationStatistics(MergerResult result)
        {
            if (result.Code == ResultCode.Error)
            {
                Console.WriteLine("Error occured: " + result.ErrorMessage);
            }
            else
            {
                Console.WriteLine("Number of merged files: " + result.NumberOfMergedFiles);
            }
        }

        /// <summary>
        /// Вывод статистики работы программы на консоль
        /// </summary>
        /// <param name="result"></param>
        public static void PrintProgramOperationStatistics(SplitterResult result)
        {
            if (result.Code == ResultCode.Error)
            {
                Console.WriteLine("Error occured: " + result.ErrorMessage);
            }
            else
            {
                Console.WriteLine("Number of files created: " + result.NumberOfResultFiles);
            }
        }

        /// <summary>
        /// Вывод статистики работы программы на консоль
        /// </summary>
        /// <param name="result"></param>
        public static void PrintProgramOperationStatistics(ColumnSplitterResult result)
        {
            if (result.Code == ResultCode.Error)
            {
                Console.WriteLine("Error occured: " + result.ErrorMessage);
            }
            else
            {
                Console.WriteLine("Number of processed rows: " + result.ProcessedRows);
                Console.WriteLine("number of columns added: " + result.CreatedColumns);
            }
        }
    }
}
