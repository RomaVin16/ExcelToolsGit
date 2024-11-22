using ClosedXML.Excel;

namespace ExcelTools.Comparison
{
    public class ComparisonWithKey
    {
        private readonly ComparisonOptions _options;
        private readonly ComparisonResult _result;

        public ComparisonWithKey(ComparisonOptions options, ComparisonResult result)
        {
            _options = options;
            _result = result;
        }

        /// <summary>
        /// Обработка листа с указанным ID 
        /// </summary>
        /// <param name="newWorkbook"></param>
        /// <param name="sourceWorksheet"></param>
        /// <param name="newWorksheet"></param>
        /// <param name="sourceFileDictionary"></param>
        /// <param name="newFileDictionary"></param>
        public void CompareByKey(XLWorkbook newWorkbook, IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet)
        {
            #region 
            var comparisonColumns = new ComparisonColumns(_options, _result);

            var sourceFileRowsDictionary = new Dictionary<string, int>();
            var newFileRowsRowsDictionary = new Dictionary<string, int>();

            ComparisonHelper.AddIds(sourceWorksheet, sourceFileRowsDictionary, _options.Id, _options.HeaderRows);
            ComparisonHelper.AddIds(newWorksheet, newFileRowsRowsDictionary, _options.Id, _options.HeaderRows);

            var sourceFileColumnsDictionary = new Dictionary<string, int>();
            var newFileColumnsDictionary = new Dictionary<string, int>();

            ComparisonHelper.AddColumnsName(sourceWorksheet, sourceFileColumnsDictionary, _options.HeaderRows);
            ComparisonHelper.AddColumnsName(newWorksheet, newFileColumnsDictionary, _options.HeaderRows);
            #endregion

            var worksheetSavingColumnsList = comparisonColumns.GetAddedColumnsNumbers(sourceFileColumnsDictionary, newFileColumnsDictionary, newWorksheet);
            comparisonColumns.CompareHeaders(newWorkbook, sourceWorksheet, newWorksheet);

            var rowCount = Math.Max(sourceWorksheet.RowsUsed().Count(), newWorksheet.RowsUsed().Count()) + _options.HeaderRows.Length;

            var deletedColumnsNumbers = comparisonColumns.GetDeletedColumnNumbers(sourceWorksheet, sourceFileColumnsDictionary, newFileColumnsDictionary);

            for (var i = sourceWorksheet.FirstRowUsed().RowNumber(); i <= rowCount; i++)
            {
                if (_options.HeaderRows != null && _options.HeaderRows.Contains(i))
                {
                    continue;
                }

                var sourceItem = ComparisonHelper.GetId(sourceWorksheet, i, _options.Id);
                var modifiedItem = ComparisonHelper.GetId(newWorksheet, i, _options.Id);

                if (newFileRowsRowsDictionary.TryGetValue(sourceItem, out var modifiedItemRowNumber))
                {
                    ProcessChangedRows(sourceWorksheet, newWorksheet, newWorkbook.Worksheet("Changed"), i, modifiedItemRowNumber, worksheetSavingColumnsList, newFileColumnsDictionary);
                    newWorkbook.Save();
                }

                if (!newFileRowsRowsDictionary.ContainsKey(sourceItem) && sourceItem != "")
                {
                    ProcessDeletedRows(sourceWorksheet, newWorksheet, newWorkbook.Worksheet("Deleted"), i, i, worksheetSavingColumnsList, deletedColumnsNumbers);
                    newWorkbook.Save();

                    newFileRowsRowsDictionary.Clear();
                    ComparisonHelper.AddIds(newWorksheet, newFileRowsRowsDictionary, _options.Id, _options.HeaderRows);
                    rowCount++;
                }

                if (!sourceFileRowsDictionary.ContainsKey(modifiedItem) && modifiedItem != "")
                {
                    newFileRowsRowsDictionary.TryGetValue(modifiedItem, out var newFileRowNumber);
                    ProcessAddedRows(sourceWorksheet, newWorksheet, newWorkbook.Worksheet("Added"), newFileRowNumber);
                    
                    newWorkbook.Save();
                }
            }
        }

        /// <summary>
        /// Вставка удаленной строки 
        /// </summary>
        /// <param name="sourceWorksheet"></param>
        /// <param name="newWorksheet"></param>
        /// <param name="deletedRows"></param>
        /// <param name="rowNumberInSourceWorksheet"></param>
        /// <param name="rowNumberInNewWorksheet"></param>
        /// <param name="worksheetSavingColumnsList"></param>
        protected void ProcessDeletedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, IXLWorksheet deletedRows, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet, List<int> worksheetSavingColumnsList, List<int> deletedColumnsNumbers)
        {
            newWorksheet.Row(rowNumberInNewWorksheet - 1).InsertRowsBelow(1);

            var rowForDeletedRowsList = 1;

            if (deletedRows.RowsUsed().Count() != 0)
            {
                rowForDeletedRowsList = deletedRows.LastRowUsed().RowNumber() + 1;
            }

            var cellsCount = sourceWorksheet.LastColumnUsed().ColumnNumber();

            var columnNumberInSourceWorksheet = sourceWorksheet.FirstColumnUsed().ColumnNumber();
            var columnNumberInNewWorksheet = sourceWorksheet.FirstColumnUsed().ColumnNumber();


            for (var j = sourceWorksheet.Row(rowNumberInSourceWorksheet).FirstCellUsed().Address.ColumnNumber; j <= cellsCount; j++)
            {
                if (worksheetSavingColumnsList.Contains(columnNumberInNewWorksheet))
                {
                    columnNumberInNewWorksheet++;
                }

                if (deletedColumnsNumbers.Contains(columnNumberInSourceWorksheet))
                {
                    columnNumberInSourceWorksheet++;
                    cellsCount--;
                }


                deletedRows.Cell(rowForDeletedRowsList, columnNumberInNewWorksheet).Value = sourceWorksheet.Cell(rowNumberInSourceWorksheet, columnNumberInSourceWorksheet).Value;
                deletedRows.Cell(rowForDeletedRowsList, columnNumberInNewWorksheet).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFF0000));

                newWorksheet.Cell(rowNumberInNewWorksheet, columnNumberInNewWorksheet).Value = sourceWorksheet.Cell(rowNumberInSourceWorksheet, columnNumberInSourceWorksheet).Value;
                newWorksheet.Cell(rowNumberInNewWorksheet, columnNumberInNewWorksheet).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFF0000));

                columnNumberInSourceWorksheet++;
                columnNumberInNewWorksheet++;
            }

            var comment = newWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().CreateComment();
            comment.AddText("Эта строка была удалена.");

            _result.CountDeletedRows++;
        }

        /// <summary>
        /// Проверка изменений в строке 
        /// </summary>
        /// <param name="sourceWorksheet"></param>
        /// <param name="newWorksheet"></param>
        /// <param name="changedRows"></param>
        /// <param name="rowNumberInSourceWorksheet"></param>
        /// <param name="rowNumberInNewWorksheet"></param>
        /// <param name="worksheetSavingColumnsList"></param>
        protected void ProcessChangedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, IXLWorksheet changedRows, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet, List<int> worksheetSavingColumnsList, Dictionary<string, int> newFileColumnsDictionary)
        {
            var rowForChangedRowsList = 1;
            var isChangedRows = false;

            if (changedRows.RowsUsed().Count() != 0)
            {
                rowForChangedRowsList = changedRows.LastRowUsed().RowNumber() + 1;
            }

            var cellsCount = sourceWorksheet.LastColumnUsed().ColumnNumber();

            var columnNumberInSourceWorksheet = sourceWorksheet.FirstColumnUsed().ColumnNumber();
            var columnNumberInNewWorksheet = sourceWorksheet.FirstColumnUsed().ColumnNumber();

            for (var i = sourceWorksheet.FirstColumnUsed().ColumnNumber(); i <= cellsCount; i++)
            {
                var sourceColumn = ComparisonHelper.GetColumnConcatenation(sourceWorksheet, _options.HeaderRows, columnNumberInSourceWorksheet);

                if (worksheetSavingColumnsList.Contains(columnNumberInNewWorksheet)) //Является ли столбец добавленным
                {
                    columnNumberInNewWorksheet++;
                    cellsCount++;
                }

                if (!newFileColumnsDictionary.ContainsKey(sourceColumn)) //Является ли столбец удаленным 
                {
                    columnNumberInSourceWorksheet++;
                }

                if (sourceWorksheet.Cell(rowNumberInSourceWorksheet, columnNumberInSourceWorksheet).Value.ToString() != newWorksheet.Cell(rowNumberInNewWorksheet, columnNumberInNewWorksheet).Value.ToString())
                {

                    newWorksheet.Cell(rowNumberInNewWorksheet, columnNumberInNewWorksheet).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFFFF00));
                    ComparisonHelper.InsertCommentInCell(newWorksheet.Cell(rowNumberInNewWorksheet, columnNumberInNewWorksheet), sourceWorksheet.Cell(rowNumberInSourceWorksheet, columnNumberInSourceWorksheet).Value.ToString(), newWorksheet.Cell(rowNumberInNewWorksheet, columnNumberInNewWorksheet).Value.ToString());

                    _result.CountChangedRows++;

                    isChangedRows = true;
                }

                columnNumberInSourceWorksheet++;
                columnNumberInNewWorksheet++;
            }

            if (isChangedRows)
            {
                var rngData = newWorksheet.Range(rowNumberInNewWorksheet, newWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().Address.ColumnNumber, rowNumberInNewWorksheet, newWorksheet.Row(rowNumberInNewWorksheet).LastCellUsed().Address.ColumnNumber);
                rngData.CopyTo(changedRows.Cell(rowForChangedRowsList, newWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().Address.ColumnNumber));
            }
        }

        /// <summary>
        /// Закрашивание добавленных строк 
        /// </summary>
        /// <param name="sourceWorksheet"></param>
        /// <param name="newWorksheet"></param>
        /// <param name="rowNumberInNewWorksheet"></param>
        protected void ProcessAddedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, IXLWorksheet addedRows, int rowNumberInNewWorksheet)
        {
            var rowForDeletedRowsList = 1;

            if (addedRows.RowsUsed().Count() != 0)
            {
                rowForDeletedRowsList = addedRows.LastRowUsed().RowNumber() + 1;
            }

            for (var j = newWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().Address.ColumnNumber; j <= newWorksheet.LastColumnUsed().ColumnNumber(); j++)
            {
                addedRows.Cell(rowForDeletedRowsList, j).Value = newWorksheet.Cell(rowNumberInNewWorksheet, j).Value;
                addedRows.Cell(rowForDeletedRowsList, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFF00B050));

                newWorksheet.Cell(rowNumberInNewWorksheet, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFF00B050));
            }

            var comment = newWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().CreateComment();
            comment.AddText("Добавлена новая строка.");

            _result.CountAddedRows++;
        }
    }
}
