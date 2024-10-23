using System.Data;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
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

            helper.AddIds(sourceWorksheet, sourceFileHash, Options.Id);
helper.AddIds(modifiedWorksheet, modifiedFileHash, Options.Id);

            var rowCount = Math.Max(sourceWorksheet.LastRowUsed().RowNumber(), modifiedWorksheet.LastRowUsed().RowNumber());

            var rowNumberInSourceWorksheet = sourceWorksheet.FirstRowUsed().RowNumber();
            var rowNumberInNewWorksheet = sourceWorksheet.FirstRowUsed().RowNumber();

            for (var i = rowNumberInSourceWorksheet; i <= rowCount + 1; i++)
            {
                var sourceItem = helper.GetId(sourceWorksheet, i, Options.Id);
                var newItem = helper.GetId(newWorksheet, i, Options.Id);

                if (!modifiedFileHash.Contains(sourceItem) && sourceItem.Length != 0)
                {
                    InsertTheDeletedRows(sourceWorksheet, newWorksheet, rowNumberInSourceWorksheet, rowNumberInNewWorksheet);
                    newWorkbook.Save();

                    rowNumberInSourceWorksheet++;
                    rowNumberInNewWorksheet++;
                    rowCount++;
                }
                else if(!sourceFileHash.Contains(newItem))
                {
                    InsertTheAddedRows(sourceWorksheet, newWorksheet, rowNumberInNewWorksheet);
                    newWorkbook.Save();

                    rowNumberInNewWorksheet++;
                }
                else
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
            foreach (var key in Options.KeysForRowsComparison)
            {
                var column = sourceWorksheet.Column(key).ColumnNumber();

                if (sourceWorksheet.Cell(rowNumberInSourceWorksheet, column).Value.ToString() != newWorksheet.Cell(rowNumberInNewWorksheet, column).Value.ToString())
                {
                    newWorksheet.Cell(rowNumberInNewWorksheet, column).Style.Fill.BackgroundColor = XLColor.Yellow;
                }
            }
        }

        protected void InsertTheAddedRows(IXLWorksheet sourceWorksheet, IXLWorksheet newWorksheet, int rowNumberInNewWorksheet)
        {
            for (var j = sourceWorksheet.Row(rowNumberInNewWorksheet).FirstCellUsed().Address.ColumnNumber; j <= newWorksheet.LastColumnUsed().ColumnNumber(); j++)
            {
                newWorksheet.Cell(rowNumberInNewWorksheet, j).Style.Fill.BackgroundColor = XLColor.Green;
            }
        }
    }
}

