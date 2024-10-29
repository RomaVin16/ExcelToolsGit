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
            var sourceFileHash = new HashSet<string>(); 
            var modifiedFileHash = new HashSet<string>();

            using var sourceWorkbook = new XLWorkbook(Options.SourceFilePath);
var sourceWorksheet = sourceWorkbook.Worksheet(Options.SheetNumber);

using var modifiedWorkbook = new XLWorkbook(Options.ModifiedFilePath);
var modifiedWorksheet = modifiedWorkbook.Worksheet(Options.SheetNumber);

using var newWorkbook = new XLWorkbook();
modifiedWorksheet.CopyTo(newWorkbook, modifiedWorksheet.Name);
var newWorksheet = newWorkbook.Worksheet(1);

            newWorkbook.SaveAs(Options.ResultFilePath);

            helper.AddIds(sourceWorksheet, sourceFileHash, Options.Id, Options.HeaderRows);
helper.AddIds(modifiedWorksheet, modifiedFileHash, Options.Id, Options.HeaderRows);

            var rowCount = Math.Max(sourceWorksheet.LastRowUsed().RowNumber(), modifiedWorksheet.LastRowUsed().RowNumber());

            var rowNumberInSourceWorksheet = sourceWorksheet.FirstRowUsed().RowNumber();
            var rowNumberInNewWorksheet = sourceWorksheet.FirstRowUsed().RowNumber();

            for (var i = rowNumberInSourceWorksheet; i <= rowCount + 1; i++)
            {
                if (Options.HeaderRows != null && Options.HeaderRows.Contains(i))
                {
                    rowNumberInSourceWorksheet++;
                    rowNumberInNewWorksheet++;
                    continue;
                }

                var sourceItem = helper.GetId(sourceWorksheet, rowNumberInSourceWorksheet, Options.Id);

                if (!helper.CheckId(sourceWorksheet, Options.Id, rowNumberInSourceWorksheet) )
                {
                    rowNumberInSourceWorksheet++;
                }

                if (!helper.CheckId(newWorksheet, Options.Id, rowNumberInNewWorksheet))
                {
                    if (!modifiedFileHash.Contains(sourceItem) && sourceItem.Length != 0)
                    {
                        InsertTheDeletedRows(sourceWorksheet, newWorksheet, rowNumberInSourceWorksheet, rowNumberInNewWorksheet);
                        newWorkbook.Save();

                        rowNumberInSourceWorksheet++;

                        rowCount++;
                    }

                    rowNumberInNewWorksheet++;
                }

                sourceItem = helper.GetId(sourceWorksheet, rowNumberInSourceWorksheet, Options.Id);
                var newItem = helper.GetId(newWorksheet, rowNumberInNewWorksheet, Options.Id);

                if (!modifiedFileHash.Contains(sourceItem) && sourceItem.Length != 0)
                {
                    InsertTheDeletedRows(sourceWorksheet, newWorksheet, rowNumberInSourceWorksheet, rowNumberInNewWorksheet);
                    newWorkbook.Save();

                    rowNumberInSourceWorksheet++;
                    rowNumberInNewWorksheet++;
                    rowCount++;
                }
                else if(!sourceFileHash.Contains(newItem) && newItem != "")
                {
                    InsertTheAddedRows(sourceWorksheet, newWorksheet, rowNumberInNewWorksheet);
                    newWorkbook.Save();

                    rowNumberInNewWorksheet++;
                    rowCount++;
                }
                else if (newItem != "" && sourceItem !="")
                {
                    CheckChanges(sourceWorksheet, newWorksheet, rowNumberInSourceWorksheet, rowNumberInNewWorksheet);
                    newWorkbook.Save();

                    rowNumberInSourceWorksheet++;
                    rowNumberInNewWorksheet++;
                }
            }
        }

        protected void InsertTheDeletedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet)
        {
            newWorksheet.Row(rowNumberInNewWorksheet - 1).InsertRowsBelow(1);

            for (var j = sourceWorksheet.Row(rowNumberInSourceWorksheet).FirstCellUsed().Address.ColumnNumber; j <= sourceWorksheet.LastColumnUsed().ColumnNumber(); j++)
            {
newWorksheet.Cell(rowNumberInNewWorksheet, j).Value = sourceWorksheet.Cell(rowNumberInSourceWorksheet, j).Value;
                newWorksheet.Cell(rowNumberInNewWorksheet, j).Style.Fill.BackgroundColor = XLColor.Red;
            }
        }

        protected void CheckChanges(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, int rowNumberInSourceWorksheet, int rowNumberInNewWorksheet)
        {
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
                }
            }
        }

        protected void InsertTheAddedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, int rowNumberInNewWorksheet)
        {
            for (var j = newWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().Address.ColumnNumber; j <= newWorksheet.LastColumnUsed().ColumnNumber(); j++)
            {
                newWorksheet.Cell(rowNumberInNewWorksheet, j).Style.Fill.BackgroundColor = XLColor.Green;
            }
        }
    }
}

