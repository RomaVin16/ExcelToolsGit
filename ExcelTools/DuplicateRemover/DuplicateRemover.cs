using ClosedXML.Excel;
using ExcelTools.Abstraction;
using ExcelTools.Exceptions;


namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemover : ExcelHandlerBase<DuplicateRemoverOptions, DuplicateRemoverResult>
    {
        /// <summary>
        /// Удаление дубликатов
        /// </summary>
        /// <param name="_options"></param>
        /// <returns></returns>
        public override DuplicateRemoverResult Process(DuplicateRemoverOptions options)
        {
            Options = options;

            if (!options.Validate())
            {
                throw new ExcelToolsException("Wrong options");
            }

            try
            {
                if (options.KeysForRowsComparison != null)
                {
                    DeleteDuplicateByKey(options.KeysForRowsComparison);
                }
                else
                {
                    DeleteDuplicateInWorkbook();
                }

                return Result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        /// <summary>
        /// Удаление дубликатов в Excel файле без ключа
        /// </summary>
        protected void DeleteDuplicateInWorkbook()
        {
            {
                using var workbook = new XLWorkbook(Options.FilePath);

                DeleteDuplicateInSheet(workbook.Worksheet(Options.SheetNumber));

                workbook.SaveAs(Options.ResultFilePath);
            }
        }

        /// <summary>
        /// Удаление дубликатов в конкретном листе
        /// </summary>
        /// <param name="item"></param>
        protected void DeleteDuplicateInSheet(IXLWorksheet item)
        {
            if (item.IsEmpty())
                return;

            var dictionary = new HashSet<string>();
            var lastRowNum = item.LastRowUsed().RowNumber();
            var i = Options.SkipRows + 1;

            while (i <= lastRowNum)
            {
                var currentRow = item.Row(i);
                i++;

                if (currentRow.IsEmpty())
                    continue;

                var currentRowKey = GetRowKey(currentRow);

                if (dictionary.Contains(currentRowKey))
                {
                    currentRow.Delete();
                    Result.RowsRemoved++;
                    lastRowNum--;
                    i--;
                }
                else
                {
                    dictionary.Add(currentRowKey);
                }
            }
        }

        /// <summary>
        /// Удаление дубликатов в листе по заданному ключу
        /// </summary>
        /// <param name="item"></param>
        /// <param name="keyColumns"></param>
        protected void DeleteDuplicateByKey(string[]? keyColumns)
        {
            using var workbook = new XLWorkbook(Options.FilePath);
            var item = workbook.Worksheet(1);

            if (item.IsEmpty())
            {
                throw new ExcelToolsException("List is empty!");
            }

            var uniqueRows = new HashSet<string>();

            var lastRowNum = item.LastRowUsed().RowNumber();
            var i = Options.SkipRows + 1;

            while (i <= lastRowNum)
            {
                var currentRow = item.Row(i);

                i++;
                if (currentRow.IsEmpty())
                    continue;

                var currentRowKey = GetRowKey(currentRow, keyColumns);

                if (uniqueRows.Contains(currentRowKey))
                {
                    currentRow.Delete();
                    Result.RowsRemoved++;
                    lastRowNum--;
                    i--;
                }
                else
                {
                    uniqueRows.Add(currentRowKey);
                }

                Result.RowsProcessed++;
            }

            workbook.SaveAs(Options.ResultFilePath);
        }

        /// <summary>
        /// Получение ключа
        /// </summary>
        /// <param name="row"></param>
        /// <param name="keyColumns"></param>
        /// <returns></returns>
        private string GetRowKey(IXLRow row, string[]? keyColumns)
        {
            return string.Join("_", keyColumns.Select(column => row.Cell(column).Value.ToString() ?? string.Empty));
        }

        /// <summary>
        /// Получение ключа строки с конкатенацией
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private string GetRowKey(IXLRow row)
        {
            return string.Join("_", row.Cells().Select(cell => cell.Value.ToString()));
        }
    }
}