using System.Diagnostics;
using ClosedXML.Excel;

namespace ExcelTools.Comparison
{
    public class ComparisonDefault
    {
        private readonly ComparisonOptions _options;
        private readonly ComparisonResult _result;

        public ComparisonDefault(ComparisonOptions options, ComparisonResult result)
        {
_options = options;
_result = result;
        }

        /// <summary>
        /// Обработка листов без указанного ID
        /// </summary>
        /// <param name="newWorkbook"></param>
        /// <param name="sourceWorksheet"></param>
        /// <param name="newWorksheet"></param>
        /// <param name="sourceFileDictionary"></param>
        /// <param name="newFileDictionary"></param>
        public void CompareWithoutKey(XLWorkbook newWorkbook, IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet)
        {
            for (var i = sourceWorksheet.FirstRowUsed().RowNumber(); i <= newWorksheet.LastRowUsed().RowNumber(); i++)
            {
                if (_options.HeaderRows != null && _options.HeaderRows.Contains(i))
                {
                    continue;
                }

                if (!sourceWorksheet.Row(i).IsEmpty() && newWorksheet.Row(i).IsEmpty())
                {
                    ProcessDeletedRows(newWorkbook, sourceWorksheet, newWorksheet, i);
                }
                else if (sourceWorksheet.Row(i).IsEmpty() && !newWorksheet.Row(i).IsEmpty())
                {
                    ProcessAddedRows(newWorkbook, newWorksheet, i);
                }
                else if (!sourceWorksheet.Row(i).IsEmpty() && !sourceWorksheet.Row(i).IsEmpty())
                {
                    ProcessChangedRows(newWorkbook, sourceWorksheet, newWorksheet, i);
                }
            }
        }

/// <summary>
/// Обработка удаленных строк  
/// </summary>
/// <param name="newWorkbook"></param>
/// <param name="sourceWorksheet"></param>
/// <param name="newWorksheet"></param>
/// <param name="rowNumber"></param>
        protected void ProcessDeletedRows(XLWorkbook newWorkbook, IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, int rowNumber)
        {
            var rowForDeletedRowsList = 1;

            if (newWorkbook.Worksheet("Deleted").RowsUsed().Count() != 0)
            {
                rowForDeletedRowsList = newWorkbook.Worksheet("Deleted").LastRowUsed().RowNumber() + 1;
            }

            for (var j = sourceWorksheet.Row(rowNumber).FirstCellUsed().Address.ColumnNumber; j <= sourceWorksheet.LastColumnUsed().ColumnNumber(); j++)
            {
                newWorkbook.Worksheet("Deleted").Cell(rowForDeletedRowsList, j).Value = sourceWorksheet.Cell(rowNumber, j).Value;
                newWorkbook.Worksheet("Deleted").Cell(rowForDeletedRowsList, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFF0000));

                newWorksheet.Cell(rowNumber, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFF0000));
            }

            _result.CountDeletedRows++;
            newWorkbook.Save();
        }

/// <summary>
/// Обработка добавленных строк 
/// </summary>
/// <param name="newWorkbook"></param>
/// <param name="newWorksheet"></param>
/// <param name="rowNumber"></param>
protected void ProcessAddedRows(XLWorkbook newWorkbook,IXLWorksheet newWorksheet, int rowNumber)
{
    var rowForDeletedRowsList = 1;

    if (newWorkbook.Worksheet("Added").RowsUsed().Count() != 0)
    {
        rowForDeletedRowsList = newWorkbook.Worksheet("Added").LastRowUsed().RowNumber() + 1;
    }

    for (var j = newWorksheet.Row(rowNumber).FirstCellUsed().Address.ColumnNumber; j <= newWorksheet.LastColumnUsed().ColumnNumber(); j++)
    {
        newWorkbook.Worksheet("Added").Cell(rowForDeletedRowsList, j).Value = newWorksheet.Cell(rowNumber, j).Value;
        newWorkbook.Worksheet("Added").Cell(rowForDeletedRowsList, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFF00B050));

        newWorksheet.Cell(rowNumber, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFF00B050));
    }

    _result.CountAddedRows++;

    newWorkbook.Save();
        }

/// <summary>
/// Обработка измененных строк 
/// </summary>
/// <param name="newWorkbook"></param>
/// <param name="sourceWorksheet"></param>
/// <param name="newWorksheet"></param>
/// <param name="rowNumber"></param>
        protected void ProcessChangedRows(XLWorkbook newWorkbook, IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, int rowNumber)
        {
            var isChageRows = false;
            var rowForChangedRowsList = 1;

            if (newWorkbook.Worksheet("Changed").RowsUsed().Count() != 0)
            {
                rowForChangedRowsList = newWorkbook.Worksheet("Changed").LastRowUsed().RowNumber() + 1;
            }

            for (var j = sourceWorksheet.FirstColumnUsed().ColumnNumber(); j <= sourceWorksheet.LastColumnUsed().ColumnNumber(); j++)
            {
                if (sourceWorksheet.Cell(rowNumber, j).Value.ToString() != newWorksheet.Cell(rowNumber, j).Value.ToString())
                {
                    newWorksheet.Cell(rowNumber, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFFFF00));
                    ComparisonHelper.InsertCommentInCell(newWorksheet.Cell(rowNumber, j), sourceWorksheet.Cell(rowNumber, j).Value.ToString(), newWorksheet.Cell(rowNumber, j).Value.ToString());

                    _result.CountChangedRows++;

                    isChageRows = true;
                }
            }

            if (isChageRows)
            {
                var rngData = newWorksheet.Range(rowNumber, newWorksheet.Row(rowNumber).FirstCellUsed().Address.ColumnNumber, rowNumber, newWorksheet.Row(rowNumber).LastCellUsed().Address.ColumnNumber);
                rngData.CopyTo(newWorkbook.Worksheet("Changed").Cell(rowForChangedRowsList, newWorksheet.Row(rowNumber).FirstCellUsed().Address.ColumnNumber));
            }
            newWorkbook.Save();
        }
    }
}
