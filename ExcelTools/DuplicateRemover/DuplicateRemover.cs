using ClosedXML.Excel;
using ExcelTools.Abstraction;

namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemover: ExcelHandlerBase<DuplicateRemoverOptions, DuplicateRemoverResult>
    {
        private DuplicateRemoverOptions options = new DuplicateRemoverOptions();
        private DuplicateRemoverResult result = new DuplicateRemoverResult();

        /// <summary>
        /// Удаление дубликатов
        /// </summary>
        /// <param name="_options"></param>
        /// <returns></returns>
        public override DuplicateRemoverResult Process(DuplicateRemoverOptions _options)
        {
            this.options = _options;

            if (!_options.Validate())
            {
                return ErrorResult("Wrong options");
            }

            try
            {
                DeleteDuplicateInWorkbook();
                return this.result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        /// <summary>
        /// Удаление дубликатов в Excel файле
        /// </summary>
        protected void DeleteDuplicateInWorkbook()
        {
            {
                using var workbook = new XLWorkbook(options.FilePath);
                foreach (var item in workbook.Worksheets)
                {
                    DeleteDuplicateInSheet(item);
                }

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
            var rowKey = currentRow.Cell(1).Value.ToString();
            dictionary.Add(rowKey, currentRow);

            for (var i = options.SkipRows + 2; i <= item.LastRowUsed().RowNumber(); i++)
            {
                result.RowsProcessed++;

                currentRow = item.Row(i);
                rowKey = currentRow.Cell(1).Value.ToString();

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
    }
}
