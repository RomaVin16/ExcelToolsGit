using ExcelTools.Abstraction;
using ExcelTools.Cleaner;
using ExcelTools.ColumnSplitter;
using ExcelTools.DuplicateRemover;
using ExcelTools.Merger;
using ExcelTools.Rotate;
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

        /// <summary>
        /// Вывод статистики работы программы на консоль
        /// </summary>
        /// <param name="result"></param>
        public static void PrintProgramOperationStatistics(RotaterResult result)
        {
            if (result.Code == ResultCode.Error)
            {
                Console.WriteLine("Error occured: " + result.ErrorMessage);
            }
            else
            {
                Console.WriteLine("Number of processed rows: " + result.RowsProcessed);
            }
        }

        public static void DeleteFolder()
        {
            var folderPath = @"C:\ExcelToolsGit\API\APIFiles";

            if (Directory.Exists(folderPath))
            {
                var subDirectories = Directory.GetDirectories(folderPath);

                foreach (var subDir in subDirectories)
                {
                    try
                    {
                        Directory.Delete(subDir, true);
                        Console.WriteLine($"Папка {subDir} успешно удалена.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при удалении папки {subDir}: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("Директория не существует.");
            }
        }
    }
}
