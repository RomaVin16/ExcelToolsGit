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
            var sourceFileDictionary = new Dictionary<string, int>();
            var newFileDictionary = new Dictionary<string, int>();

            ComparisonHelper.AddIds(sourceWorksheet, sourceFileDictionary, _options.Id, _options.HeaderRows);
            ComparisonHelper.AddIds(newWorksheet, newFileDictionary, _options.Id, _options.HeaderRows);

            var rowCount = Math.Max(sourceWorksheet.RowsUsed().Count(), newWorksheet.RowsUsed().Count());

            for (var i = sourceWorksheet.FirstRowUsed().RowNumber(); i <= rowCount; i++)
            {
                if (_options.HeaderRows != null && _options.HeaderRows.Contains(i))
                {
                    continue;
                }

                var sourceItem = ComparisonHelper.GetId(sourceWorksheet, i, _options.Id);
                var modifiedItem = ComparisonHelper.GetId(newWorksheet, i, _options.Id);

                if (newFileDictionary.TryGetValue(sourceItem, out var modifiedItemRowNumber))
                {
                    ProcessChangedRows(sourceWorksheet, newWorksheet, newWorkbook.Worksheet("Changed"), i, modifiedItemRowNumber);
                    newWorkbook.Save();
                }

                if (!newFileDictionary.ContainsKey(sourceItem) && sourceItem != "")
                {
                    ProcessDeletedRows(sourceWorksheet, newWorksheet, newWorkbook.Worksheet("Deleted"), i, i);
                    newWorkbook.Save();

                    newFileDictionary.Clear();
                    ComparisonHelper.AddIds(newWorksheet, newFileDictionary, _options.Id, _options.HeaderRows);
                    rowCount++;
                }

                if (!sourceFileDictionary.ContainsKey(modifiedItem) && modifiedItem != "")
                {
                    newFileDictionary.TryGetValue(modifiedItem, out var newFileRowNumber);
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
        /// <param name="rowNumberInSourceWorksheet"></param>
        /// <param name="rowNumberInNewWorksheet"></param>
        protected void ProcessDeletedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, IXLWorksheet deletedRows, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet)
        {
            newWorksheet.Row(rowNumberInNewWorksheet - 1).InsertRowsBelow(1);

            var rowForDeletedRowsList = 1;

            if (deletedRows.RowsUsed().Count() != 0)
            {
                rowForDeletedRowsList = deletedRows.LastRowUsed().RowNumber() + 1;
            }

            for (var j = sourceWorksheet.Row(rowNumberInSourceWorksheet).FirstCellUsed().Address.ColumnNumber; j <= sourceWorksheet.LastColumnUsed().ColumnNumber(); j++)
            {
                deletedRows.Cell(rowForDeletedRowsList, j).Value = sourceWorksheet.Cell(rowNumberInSourceWorksheet, j).Value;
                deletedRows.Cell(rowForDeletedRowsList, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFF0000));

                newWorksheet.Cell(rowNumberInNewWorksheet, j).Value = sourceWorksheet.Cell(rowNumberInSourceWorksheet, j).Value;
                newWorksheet.Cell(rowNumberInNewWorksheet, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFF0000));
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
        /// <param name="rowNumberInSourceWorksheet"></param>
        /// <param name="rowNumberInNewWorksheet"></param>
        protected void ProcessChangedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, IXLWorksheet changedRows, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet)
        {
            var rowForChangedRowsList = 1;
            var isChangedRows = false;

            if (changedRows.RowsUsed().Count() != 0)
            {
                rowForChangedRowsList = changedRows.LastRowUsed().RowNumber() + 1;
            }

            for (var i = sourceWorksheet.FirstColumnUsed().ColumnNumber(); i <= sourceWorksheet.LastColumnUsed().ColumnNumber(); i++)
            {
                var columnName = sourceWorksheet.Column(i).ColumnLetter();

                if (_options.Id.Contains(columnName))
                {
                    continue;
                }

                if (sourceWorksheet.Cell(rowNumberInSourceWorksheet, i).Value.ToString() != newWorksheet.Cell(rowNumberInNewWorksheet, i).Value.ToString())
                {

                    newWorksheet.Cell(rowNumberInNewWorksheet, i).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFFFF00));
                    ComparisonHelper.InsertCommentInCell(newWorksheet.Cell(rowNumberInNewWorksheet, i), sourceWorksheet.Cell(rowNumberInSourceWorksheet, i).Value.ToString(), newWorksheet.Cell(rowNumberInNewWorksheet, i).Value.ToString());

                    _result.CountChangedRows++;

                    isChangedRows = true;
                }
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
