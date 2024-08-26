using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTools.Abstraction;


namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemover : ExcelHandlerBase<DuplicateRemoverOptions, DuplicateRemoverResult>
    {
        private DuplicateRemoverOptions options = new DuplicateRemoverOptions();
        private DuplicateRemoverResult result = new DuplicateRemoverResult();

        /// <summary>
        /// Удаление дубликатов
        /// </summary>
        /// <param name="_options"></param>
        /// <returns></returns>
        public override DuplicateRemoverResult Process(DuplicateRemoverOptions options)
        {
            this.options = options;

            if (!options.Validate())
            {
                return ErrorResult("Wrong options");
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

                return this.result;
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
                using var workbook = new XLWorkbook(options.FilePath);

                DeleteDuplicateInSheet(workbook.Worksheet(1));

                workbook.SaveAs(options.ResultFilePath);
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

            var dictionary = new Dictionary<string, object?>();

            for (var i = options.SkipRows + 1; i <= item.LastRowUsed().RowNumber(); i++)
            {
                result.RowsProcessed++;

                var currentRow = item.Row(i);

                if (currentRow.IsEmpty())
                    continue;

                var currentRowKey = GetRowKey(currentRow);

                if (i == 1)
                {
                    currentRow = item.Row(1);
                    currentRowKey = GetRowKey(currentRow);
                    dictionary.Add(currentRowKey, null);
                }

                if (dictionary.ContainsKey(currentRowKey))
                {
                    currentRow.Delete();

                    result.RowsRemoved++;
                }
                else
                {
                    dictionary.Add(currentRowKey, currentRow);
                }
            }
        }

        /// <summary>
        /// Удаление дубликатов в листе по заданному ключу
        /// </summary>
        /// <param name="item"></param>
        /// <param name="keyColumns"></param>
        protected void DeleteDuplicateByKey(string[] keyColumns)
        {
            using var workbook = new XLWorkbook(options.FilePath);
            var item = workbook.Worksheet(2);

            if (item.IsEmpty())
            {
                throw new Exception("List is empty.");
            }

            var uniqueRows = new HashSet<string>();

            for (var i = options.SkipRows + 1; i <= item.LastRowUsed().RowNumber() + 1; i++)
            {
                result.RowsProcessed++;

                var currentRow = item.Row(i);

                if (currentRow.IsEmpty())
                    continue;

                var currentRowKey = GetRowKey(currentRow, keyColumns);

                if (i == 1)
                {
                    var firstRow = item.FirstRowUsed();
                    var firstRowKey = GetRowKey(firstRow, keyColumns);
                    uniqueRows.Add(firstRowKey);
                    continue;
                }

                if (uniqueRows.Contains(currentRowKey))
                {
                    currentRow.Delete();
                    result.RowsRemoved++;
                }
                else
                {
                    uniqueRows.Add(currentRowKey);
                }
            }

            workbook.SaveAs(options.ResultFilePath);
        }

        /// <summary>
        /// Получение ключа
        /// </summary>
        /// <param name="row"></param>
        /// <param name="keyColumns"></param>
        /// <returns></returns>
        private string GetRowKey(IXLRow row, string[] keyColumns)
        {
            return string.Join(",", keyColumns.Select(column => row.Cell(column).Value.ToString() ?? string.Empty));
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