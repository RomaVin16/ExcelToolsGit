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
            var helper = new ComparisonHelper();
            var sourceFileDictionary = new Dictionary<string, int>();
            var newFileDictionary = new Dictionary<string, int>();

            using var sourceWorkbook = new XLWorkbook(Options.SourceFilePath);
var sourceWorksheet = sourceWorkbook.Worksheet(Options.SheetNumber);

using var modifiedWorkbook = new XLWorkbook(Options.ModifiedFilePath);
var modifiedWorksheet = modifiedWorkbook.Worksheet(Options.SheetNumber);

using var newWorkbook = new XLWorkbook();
modifiedWorksheet.CopyTo(newWorkbook, modifiedWorksheet.Name);
var newWorksheet = newWorkbook.Worksheet(1);

            newWorkbook.SaveAs(Options.ResultFilePath);

            helper.AddIds(sourceWorksheet, sourceFileDictionary, Options.Id, Options.HeaderRows);
helper.AddIds(newWorksheet, newFileDictionary, Options.Id, Options.HeaderRows);

var rowCount = Math.Max(sourceWorksheet.RowsUsed().Count(), newWorksheet.RowsUsed().Count());

for (var i = sourceWorksheet.FirstRowUsed().RowNumber(); i <= rowCount; i++)
{
    if (Options.HeaderRows != null && Options.HeaderRows.Contains(i))
    {
        continue;
    }

                var sourceItem = helper.GetId(sourceWorksheet, i, Options.Id);
    var modifiedItem = helper.GetId(newWorksheet, i, Options.Id);

    if (newFileDictionary.TryGetValue(sourceItem, out var modifiedItemRowNumber))
    {
        CheckChanges(sourceWorksheet, newWorksheet, i, modifiedItemRowNumber);
        newWorkbook.Save();
    }

    if (!newFileDictionary.ContainsKey(sourceItem) && sourceItem != "")
    {
        InsertTheDeletedRows(sourceWorksheet, newWorksheet, i, i);
        newWorkbook.Save();
        newFileDictionary.Clear();
        helper.AddIds(newWorksheet, newFileDictionary, Options.Id, Options.HeaderRows);
        rowCount++;
    }

                if (!sourceFileDictionary.ContainsKey(modifiedItem) && modifiedItem != "")
                {
                     newFileDictionary.TryGetValue(modifiedItem, out var newFileRowNumber);
                    InsertTheAddedRows(sourceWorksheet, newWorksheet, newFileRowNumber);
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
        protected void InsertTheDeletedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet)
        {
            newWorksheet.Row(rowNumberInNewWorksheet - 1).InsertRowsBelow(1);

            for (var j = sourceWorksheet.Row(rowNumberInSourceWorksheet).FirstCellUsed().Address.ColumnNumber; j <= sourceWorksheet.LastColumnUsed().ColumnNumber(); j++)
            {
newWorksheet.Cell(rowNumberInNewWorksheet, j).Value = sourceWorksheet.Cell(rowNumberInSourceWorksheet, j).Value;
                newWorksheet.Cell(rowNumberInNewWorksheet, j).Style.Fill.BackgroundColor = XLColor.CadmiumRed;
            }

            var comment = newWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().CreateComment();
            comment.AddText("Это строка была удалена.");
        }

/// <summary>
/// Проверка изменений в строке 
/// </summary>
/// <param name="sourceWorksheet"></param>
/// <param name="newWorksheet"></param>
/// <param name="rowNumberInSourceWorksheet"></param>
/// <param name="rowNumberInNewWorksheet"></param>
        protected void CheckChanges(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet)
{
    var helper = new ComparisonHelper();

            for (var i = sourceWorksheet.FirstColumnUsed().ColumnNumber(); i <= sourceWorksheet.LastColumnUsed().ColumnNumber(); i++)
            {
                var columnName = sourceWorksheet.Column(i).ColumnLetter();

                if (Options.Id.Contains(columnName))
                {
                    continue;
                }

                if (sourceWorksheet.Cell(rowNumberInSourceWorksheet, i).Value.ToString() != newWorksheet.Cell(rowNumberInNewWorksheet,i).Value.ToString())
                {
                    newWorksheet.Cell(rowNumberInNewWorksheet, i).Style.Fill.BackgroundColor = XLColor.Yellow;

                    helper.InsertCommentInCell(newWorksheet.Cell(rowNumberInNewWorksheet, i), sourceWorksheet.Cell(rowNumberInSourceWorksheet, i).Value.ToString(), newWorksheet.Cell(rowNumberInNewWorksheet, i).Value.ToString());
                }
            }
}

/// <summary>
/// Закрашивание добавленных строк 
/// </summary>
/// <param name="sourceWorksheet"></param>
/// <param name="newWorksheet"></param>
/// <param name="rowNumberInNewWorksheet"></param>
protected void InsertTheAddedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, int rowNumberInNewWorksheet)
{
    for (var j = newWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().Address.ColumnNumber; j <= newWorksheet.LastColumnUsed().ColumnNumber(); j++)
    {
        newWorksheet.Cell(rowNumberInNewWorksheet, j).Style.Fill.BackgroundColor = XLColor.Green;
    }

    var comment = newWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().CreateComment();
    comment.AddText("Добавлена новая строка.");
}
    }
}

