using ClosedXML.Excel;

namespace ExcelTools.Comparison
{
    public class ComparisonColumns
    {
        private readonly ComparisonOptions _options;
        private readonly ComparisonResult _result;

        public ComparisonColumns(ComparisonOptions options, ComparisonResult result)
        {
            _options = options;
            _result = result;
        }

        public void CompareHeaders(XLWorkbook newWorkbook, IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet)
        {
            var sourceFileColumnsDictionary = new Dictionary<string, int>();
            var newFileColumnsDictionary = new Dictionary<string, int>();

            ComparisonHelper.AddColumnsName(sourceWorksheet, sourceFileColumnsDictionary, _options.HeaderRows);
            ComparisonHelper.AddColumnsName(newWorksheet, newFileColumnsDictionary, _options.HeaderRows);


            newWorkbook.Save();

            var columnsCount = Math.Max(sourceWorksheet.ColumnsUsed().Count(), newWorksheet.ColumnsUsed().Count());

            for (var i = sourceWorksheet.FirstColumnUsed().ColumnNumber(); i <= columnsCount; i++)
            {
                var modifiedItem = ComparisonHelper.GetColumnConcatenation(newWorksheet, _options.HeaderRows, i);

                if (!sourceFileColumnsDictionary.ContainsKey(modifiedItem) && modifiedItem != "_")  //обработка нового столбца
                {
                    for (var j = sourceWorksheet.Column(i).FirstCellUsed().Address.RowNumber; j <= sourceWorksheet.LastRowUsed().RowNumber(); j++)
                    {
                        newWorksheet.Cell(j, i).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFF00B050));
                    }
                }

                newWorkbook.Save();
            }

        }

        /// <summary>
        /// Получение номеров новых столбцов 
        /// </summary>
        public List<int> GetAddedColumnsNumbers(Dictionary<string, int> sourceFileColumnsDictionary, Dictionary<string, int> newFileColumnsDictionary,  IXLWorksheet newWorksheet)
        {
            var worksheetSavingColumnsList = new List<int>();

                        var columnsCount = newWorksheet.ColumnsUsed().Count();

            for (var i = newWorksheet.FirstColumnUsed().ColumnNumber(); i <= columnsCount; i++)
            {
                var modifiedItem = ComparisonHelper.GetColumnConcatenation(newWorksheet, _options.HeaderRows, i);

                if (!sourceFileColumnsDictionary.ContainsKey(modifiedItem) && modifiedItem != "_")
                {
                    worksheetSavingColumnsList.Add(i);

                }
            }

            return worksheetSavingColumnsList;
        }

        /// <summary>
        /// Получение номеров удаленных столбцов 
        /// </summary>
        /// <param name="sourceWorksheet"></param>
        /// <param name="newWorksheet"></param>
        /// <param name="sourceFileColumnsDictionary"></param>
        /// <param name="newFileColumnsDictionary"></param>
        /// <param name="sourceFileRowsDictionary"></param>
        /// <param name="newFileRowsDictionary"></param>
        /// <returns></returns>
        public List<int> GetDeletedColumnNumbers(IXLWorksheet sourceWorksheet, Dictionary<string, int> sourceFileColumnsDictionary, Dictionary<string, int> newFileColumnsDictionary)
        {
            var deletedColumnsNumbers = new List<int>();
            
            for (var i = sourceWorksheet.FirstColumnUsed().ColumnNumber(); i <= sourceWorksheet.ColumnsUsed().Count(); i++)
            {
                var sourceItem = ComparisonHelper.GetColumnConcatenation(sourceWorksheet, _options.HeaderRows, i);

                sourceFileColumnsDictionary.TryGetValue(sourceItem, out var columnNumberInSourceWorksheet);

                if (!newFileColumnsDictionary.ContainsKey(sourceItem) && sourceItem != "_") 
                {
                    deletedColumnsNumbers.Add(columnNumberInSourceWorksheet);
                }
            }

            return deletedColumnsNumbers;
        }

        /// <summary>
        /// Вставка удаленного столбца   
        /// </summary>
        /// <param name="sourceWorksheet"></param>
        /// <param name="newWorksheet"></param>
        /// <param name="columnNumberInSourceWorksheet"></param>
        /// <param name="columnNumberInNewWorksheet"></param>
        public void ProcessDeletedColumn(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, Dictionary<string, int> sourceFileColumnsDictionary, Dictionary<string, int> newFileColumnsDictionary, Dictionary<string, int> sourceFileRowsDictionary, Dictionary<string,int> newFileRowsDictionary)
        {
            var columnsCount = Math.Max(sourceWorksheet.ColumnsUsed().Count(), newWorksheet.ColumnsUsed().Count());

            for (var i = sourceWorksheet.FirstColumnUsed().ColumnNumber(); i <= columnsCount; i++)
            {
                var sourceItem = ComparisonHelper.GetColumnConcatenation(sourceWorksheet, _options.HeaderRows, i);

                    sourceFileColumnsDictionary.TryGetValue(sourceItem, out var columnNumberInSourceWorksheet);
                    var columnNumberInNewWorksheet = newWorksheet.LastColumnUsed().ColumnNumber() + 1;

                if (!newFileColumnsDictionary.ContainsKey(sourceItem) && sourceItem != "_") //обработка удаленного столбца
                {
                    for (var j = sourceWorksheet.Column(columnNumberInSourceWorksheet).FirstCellUsed().Address.RowNumber; j <= sourceWorksheet.LastRowUsed().RowNumber(); j++)
                    {
                        newWorksheet.Cell(j, columnNumberInNewWorksheet).Value = sourceWorksheet.Cell(j, columnNumberInSourceWorksheet).Value;
                        newWorksheet.Cell(j, columnNumberInNewWorksheet).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFF0000));
                    }

                    _result.CountDeletedColumns++;

                    columnsCount++;
                }
            }
        }
    }
}
