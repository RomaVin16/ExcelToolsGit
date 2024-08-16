using ClosedXML.Excel;
using ExcelTools.Cleaner;

namespace ExcelTools.Helpers
{
    public static class ExcelHelper
    {
        /// <summary>
        /// Проверка файла на пустоту
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        /// <exception cref="Exception"></exception>
        public static bool CheckingTheFileForEmptiness(IXLWorksheet item, CleanResult result)
        {
            return item.IsEmpty();
        }
    }
}