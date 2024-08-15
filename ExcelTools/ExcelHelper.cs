using ClosedXML.Excel;

namespace ExcelTools
{
    public static class ExcelHelper
    {
        /// <summary>
        /// Проверка файла на пустоту
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        /// <exception cref="Exception"></exception>
        public static void CheckingTheFileForEmptiness(IXLWorksheet item, CleanResult result)
        {
            if (item.IsEmpty())
            {
                throw new Exception("Переданный для обработки файл пуст.");
            }
        }

        /// <summary>
        /// Проверка расширения файла
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static XLWorkbook WriteFile(string fileName, CleanResult result)
        {
            XLWorkbook wordbook;

            try
            {
                wordbook = new XLWorkbook(fileName);
            }
            catch (Exception)
            {
                throw new Exception("Переданный файл не найден либо имеет некорректное расширение.");
            }

            return wordbook;
        }

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
