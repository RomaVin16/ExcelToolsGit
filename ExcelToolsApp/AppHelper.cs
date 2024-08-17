using ExcelTools.Abstraction;
using ExcelTools.Cleaner;

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
    }
}
