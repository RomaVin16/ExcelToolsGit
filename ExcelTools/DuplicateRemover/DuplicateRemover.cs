using ClosedXML.Excel;
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
                foreach (var item in workbook.Worksheets)
                {
                    DeleteDuplicateInSheet(item);
                }

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

            var dictionary = new Dictionary<string, object>();

            var currentRow = item.FirstRowUsed();
            var rowKey = GetRowKey(currentRow);
            dictionary.Add(rowKey, currentRow);

            for (var i = options.SkipRows + 2; i <= item.LastRowUsed().RowNumber(); i++)
            {
                result.RowsProcessed++;

                currentRow = item.Row(i);
                rowKey = GetRowKey(currentRow);

                if (currentRow.IsEmpty())
                    continue;

                if (dictionary.TryGetValue(rowKey, out var value))
                {

                    var stringBeingChecked = value as IXLRow;

                    if (!CompareRows(currentRow, stringBeingChecked)) continue;
                    currentRow.Delete();
                    result.RowsRemoved++;
                }
                else
                {
                    dictionary.Add(rowKey, currentRow);
                }
            }
        }

        /// <summary>
        /// Проверка строк на дублирование
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="stringBeingChecked"></param>
        /// <returns></returns>
        protected bool CompareRows(IXLRow currentRow, IXLRow stringBeingChecked)
        {
            var isDuplicate = true;

            for (var i = 2; i <= currentRow.LastCellUsed().Address.ColumnNumber; i++)
            {
                var currentCell = currentRow.Cell(i);
                var cellBeingChecked = stringBeingChecked?.Cell(i);

                if (currentCell.Value.ToString() != cellBeingChecked.Value.ToString() | currentCell.FormulaA1 != cellBeingChecked.FormulaA1)
                {
                    return false;
                }
            }

            return isDuplicate;
        }

/// <summary>
/// Удаление дубликатов в листе по заданному ключу
/// </summary>
/// <param name="item"></param>
/// <param name="keyColumns"></param>
        protected void DeleteDuplicateByKey(string[] keyColumns)
        {
            using var workbook = new XLWorkbook(options.FilePath);
            var item = workbook.Worksheet(1);

            var uniqueRows = new HashSet<string>();

            var firstRow = item.FirstRowUsed();
            var firstRowKey = GetRowKey(firstRow, keyColumns);
            uniqueRows.Add(firstRowKey);

            for (var i = options.SkipRows + 2; i <= item.LastRowUsed().RowNumber(); i++)
            {
                result.RowsProcessed++;

                var currentRow = item.Row(i);

                if (currentRow.IsEmpty())
                    continue;

                var currentRowKey = GetRowKey(currentRow, keyColumns);

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
            return string.Join("_", row.Cells().Select(cell => cell.Value.ToString() ?? string.Empty));
        }
    }
}