using ClosedXML.Excel;
using ExcelTools.Abstraction;
using ExcelTools.Exceptions;

namespace ExcelTools.Comparison
{
    public class Comparison: ExcelHandlerBase<ComparisonOptions, ComparisonResult>
    {
        public override ComparisonResult Process(ComparisonOptions options)
        {
            Options = options;

            if (!options.Validate())
            {
                throw new ExcelToolsException("Wrong options");
            }

            try
            {
                Compare();
                return Result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        protected void Compare()
        {
            var sourceFileDictionary = new Dictionary<string, int>();
            var newFileDictionary = new Dictionary<string, int>();

            using var sourceWorkbook = new XLWorkbook(Options.SourceFilePath);
var sourceWorksheet = sourceWorkbook.Worksheet(Options.SheetNumber);

using var modifiedWorkbook = new XLWorkbook(Options.ModifiedFilePath);
var modifiedWorksheet = modifiedWorkbook.Worksheet(Options.SheetNumber);

using var newWorkbook = new XLWorkbook();
modifiedWorksheet.CopyTo(newWorkbook, modifiedWorksheet.Name);
var newWorksheet = newWorkbook.Worksheet(1);

newWorkbook.AddWorksheet("Changed");
newWorkbook.AddWorksheet("Deleted");
newWorkbook.AddWorksheet("Added");

ComparisonHelper.InsertHeadersInList(sourceWorksheet, newWorkbook, Options.HeaderRows);

newWorkbook.SaveAs(Options.ResultFilePath);

ComparisonHelper.AddIds(sourceWorksheet, sourceFileDictionary, Options.Id, Options.HeaderRows);
ComparisonHelper.AddIds(newWorksheet, newFileDictionary, Options.Id, Options.HeaderRows);

if (Options.Id != null)
{
    ProcessWorksheetWithId(newWorkbook, sourceWorksheet, newWorksheet, sourceFileDictionary, newFileDictionary);
}
else
{
    ProcessWorksheetWithoutId(newWorkbook, sourceWorksheet, newWorksheet, sourceFileDictionary, newFileDictionary);
}

        }

        /// <summary>
        /// Обработка листа с указанным ID 
        /// </summary>
        /// <param name="newWorkbook"></param>
        /// <param name="sourceWorksheet"></param>
        /// <param name="newWorksheet"></param>
        /// <param name="sourceFileDictionary"></param>
        /// <param name="newFileDictionary"></param>
        protected void ProcessWorksheetWithId(XLWorkbook newWorkbook, IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, Dictionary<string, int> sourceFileDictionary, Dictionary<string, int> newFileDictionary)
        {
            var rowCount = Math.Max(sourceWorksheet.RowsUsed().Count(), newWorksheet.RowsUsed().Count());

            for (var i = sourceWorksheet.FirstRowUsed().RowNumber(); i <= rowCount; i++)
            {
                if (Options.HeaderRows != null && Options.HeaderRows.Contains(i))
                {
                    continue;
                }

                var sourceItem = ComparisonHelper.GetId(sourceWorksheet, i, Options.Id);
                var modifiedItem = ComparisonHelper.GetId(newWorksheet, i, Options.Id);

                if (newFileDictionary.TryGetValue(sourceItem, out var modifiedItemRowNumber))
                {
                    CheckChanges(sourceWorksheet, newWorksheet, newWorkbook.Worksheet("Changed"), i, modifiedItemRowNumber);
                    newWorkbook.Save();
                }

                if (!newFileDictionary.ContainsKey(sourceItem) && sourceItem != "")
                {
                    InsertTheDeletedRows(sourceWorksheet, newWorksheet, newWorkbook.Worksheet("Deleted"), i, i);
                    newWorkbook.Save();
                    newFileDictionary.Clear();
                    ComparisonHelper.AddIds(newWorksheet, newFileDictionary, Options.Id, Options.HeaderRows);
                    rowCount++;
                }

                if (!sourceFileDictionary.ContainsKey(modifiedItem) && modifiedItem != "")
                {
                    newFileDictionary.TryGetValue(modifiedItem, out var newFileRowNumber);
                    InsertTheAddedRows(sourceWorksheet, newWorksheet, newWorkbook.Worksheet("Added"), newFileRowNumber);
                    newWorkbook.Save();
                }
            }
        }

/// <summary>
/// Обработка листов без указанного ID
/// </summary>
/// <param name="newWorkbook"></param>
/// <param name="sourceWorksheet"></param>
/// <param name="newWorksheet"></param>
/// <param name="sourceFileDictionary"></param>
/// <param name="newFileDictionary"></param>
        protected void ProcessWorksheetWithoutId(XLWorkbook newWorkbook, IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, Dictionary<string, int> sourceFileDictionary, Dictionary<string, int> newFileDictionary)
        {
            for (var i = sourceWorksheet.FirstRowUsed().RowNumber(); i <= newWorksheet.LastRowUsed().RowNumber(); i++)
            {
                if (Options.HeaderRows != null && Options.HeaderRows.Contains(i))
                {
                    continue;
                }

                if (!sourceWorksheet.Row(i).IsEmpty() && newWorksheet.Row(i).IsEmpty())
                {
                    var rowForDeletedRowsList = 1;

                    if (newWorkbook.Worksheet("Deleted").RowsUsed().Count() != 0)
                    {
                        rowForDeletedRowsList = newWorkbook.Worksheet("Deleted").LastRowUsed().RowNumber() + 1;
                    }

                    for (var j = sourceWorksheet.Row(i).FirstCellUsed().Address.ColumnNumber; j <= sourceWorksheet.LastColumnUsed().ColumnNumber(); j++)
                    {
                        newWorkbook.Worksheet("Deleted").Cell(rowForDeletedRowsList, j).Value = sourceWorksheet.Cell(i, j).Value;
                        newWorkbook.Worksheet("Deleted").Cell(rowForDeletedRowsList, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFF0000));

                        newWorksheet.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFF0000));
                    }
                    newWorkbook.Save();
                }
                else if (sourceWorksheet.Row(i).IsEmpty() && !newWorksheet.Row(i).IsEmpty())
                {
                    var rowForDeletedRowsList = 1;

                    if (newWorkbook.Worksheet("Added").RowsUsed().Count() != 0)
                    {
                        rowForDeletedRowsList = newWorkbook.Worksheet("Added").LastRowUsed().RowNumber() + 1;
                    }

                    for (var j = newWorksheet.Row(i).FirstCellUsed().Address.ColumnNumber; j <= newWorksheet.LastColumnUsed().ColumnNumber(); j++)
                    {
                        newWorkbook.Worksheet("Added").Cell(rowForDeletedRowsList, j).Value = newWorksheet.Cell(i, j).Value;
                        newWorkbook.Worksheet("Added").Cell(rowForDeletedRowsList, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFF00B050));

                        newWorksheet.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFF00B050));
                    }
                    newWorkbook.Save();
                }
                else if (!sourceWorksheet.Row(i).IsEmpty() && !sourceWorksheet.Row(i).IsEmpty())
                {
                    var isChageRows = false;
                    var rowForChangedRowsList = 1;

                    if (newWorkbook.Worksheet("Changed").RowsUsed().Count() != 0)
                    {
                        rowForChangedRowsList = newWorkbook.Worksheet("Changed").LastRowUsed().RowNumber() + 1;
                    }

                    for (var j = sourceWorksheet.FirstColumnUsed().ColumnNumber(); j <= sourceWorksheet.LastColumnUsed().ColumnNumber(); j++)
                    {
                        if (sourceWorksheet.Cell(i, j).Value.ToString() != newWorksheet.Cell(i, j).Value.ToString())
                        {
                            newWorksheet.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFFFF00));
                            isChageRows = true;
                        }
                    }

                    if (isChageRows)
                    {
                        var rngData = newWorksheet.Range(i, newWorksheet.Row(i).FirstCellUsed().Address.ColumnNumber, i, newWorksheet.Row(i).LastCellUsed().Address.ColumnNumber);
                        rngData.CopyTo(newWorkbook.Worksheet("Changed").Cell(rowForChangedRowsList, newWorksheet.Row(i).FirstCellUsed().Address.ColumnNumber));
                    }
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
        protected void InsertTheDeletedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, IXLWorksheet deletedRows, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet)
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
        }

/// <summary>
/// Проверка изменений в строке 
/// </summary>
/// <param name="sourceWorksheet"></param>
/// <param name="newWorksheet"></param>
/// <param name="rowNumberInSourceWorksheet"></param>
/// <param name="rowNumberInNewWorksheet"></param>
        protected void CheckChanges(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, IXLWorksheet changedRows, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet)
{ 
    var rowForChangedRowsList = 1;

    var isChageRows = false;


            if (changedRows.RowsUsed().Count() != 0)
    {
        rowForChangedRowsList = changedRows.LastRowUsed().RowNumber() + 1;
    }

    for (var i = sourceWorksheet.FirstColumnUsed().ColumnNumber(); i <= sourceWorksheet.LastColumnUsed().ColumnNumber(); i++)
    {
        var columnName = sourceWorksheet.Column(i).ColumnLetter();


        if (Options.Id.Contains(columnName))
        {
            continue;
        }

        if (sourceWorksheet.Cell(rowNumberInSourceWorksheet, i).Value.ToString() != newWorksheet.Cell(rowNumberInNewWorksheet,i).Value.ToString())
        {

            newWorksheet.Cell(rowNumberInNewWorksheet, i).Style.Fill.BackgroundColor = XLColor.FromArgb(unchecked((int)0xFFFFFF00));
                    ComparisonHelper.InsertCommentInCell(newWorksheet.Cell(rowNumberInNewWorksheet, i), sourceWorksheet.Cell(rowNumberInSourceWorksheet, i).Value.ToString(), newWorksheet.Cell(rowNumberInNewWorksheet, i).Value.ToString());

            isChageRows = true;
        }
    }

    if (isChageRows)
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
protected void InsertTheAddedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, IXLWorksheet addedRows, int rowNumberInNewWorksheet)
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
}
    }
}

