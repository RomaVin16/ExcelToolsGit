using ExcelTools.Abstraction;
using ExcelTools.Cleaner;
using ExcelTools.DuplicateRemover;
using ExcelTools.Merger;

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
                Console.WriteLine("Program is success");
            }

        }
    }
}
