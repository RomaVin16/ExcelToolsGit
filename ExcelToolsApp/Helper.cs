using ExcelTools.Cleaner;

namespace ExcelTools.App
{
    public static class Helper
    {
        /// <summary>
        /// Вывод статистики работы программы на консоль
        /// </summary>
        /// <param name="result"></param>
        public static void PrintProgramOperationStatistics(CleanResult result)
        {
            if (result.Code == (CleanResult.ResultCode)1)
            {
                Console.WriteLine("В ходе работы программы произошла ошибка: " + result.ErrorMessage);
            }
            else
            {
                Console.WriteLine("Кол-во обработанных строк: " + result.RowsProcessed);
                Console.WriteLine("Кол-во удаленных строк: " + result.RowsRemoved);
            }
        }
    }
}
