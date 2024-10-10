using ClosedXML.Excel;
using ExcelTools.Abstraction;
using ExcelTools.Exceptions;

namespace ExcelTools.Rotate
{
    public class Rotater: ExcelHandlerBase<RotaterOptions, RotaterResult>
    {
        public override RotaterResult Process(RotaterOptions options)
        {
            Options = options;

            if (!options.Validate())
            {
                throw new ExcelToolsException("Wrong options");
            }

            try
            {
                RotateTheTable();
                return Result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        protected void RotateTheTable()
        {
            using var workbook = new XLWorkbook(Options.FilePath);
            using var newWorkbook = new XLWorkbook();
            var newWorksheet = newWorkbook.AddWorksheet();

            var worksheet = workbook.Worksheet(Options.SheetNumber);

            var outputSheet = workbook.AddWorksheet();

            var rowCount = worksheet.LastRowUsed().RowNumber();
            var colCount = worksheet.LastColumnUsed().ColumnNumber();

            for (var i = worksheet.FirstRowUsed().RowNumber(); i <= rowCount; i++)
            {
                for (var j = worksheet.FirstColumnUsed().ColumnNumber(); j <= colCount; j++)
                {
                    if (worksheet.Cell(i, j).IsMerged())
                    {
                        var mergedRange = worksheet.MergedRanges.FirstOrDefault(r => r.Contains(worksheet.Cell(i, j)));

                        if (mergedRange != null)
                        {
                            var value = worksheet.Cell(mergedRange.FirstRow().RowNumber(), mergedRange.FirstColumn().ColumnNumber()).Value;

                            var firstRow = mergedRange.FirstRow().RowNumber();
                            var lastRow = mergedRange.LastRow().RowNumber();
                            var firstColumn = mergedRange.FirstColumn().ColumnNumber();
                            var lastColumn = mergedRange.LastColumn().ColumnNumber();

                            var newMergedRange = outputSheet.Range(outputSheet.Cell(firstColumn, firstRow), outputSheet.Cell(lastColumn, lastRow));
                            newMergedRange.Merge();

                            outputSheet.Cell(j, i).Value = value;
                        }
                    }
                    else
                    {
                        var value = worksheet.Cell(i, j).Value;
                        outputSheet.Cell(j, i).Value = value;
                    }
                }
            }

            var firstTableCell = outputSheet.Cell(outputSheet.FirstRowUsed().RowNumber() + Options.SkipRows,
                outputSheet.FirstColumnUsed().ColumnNumber());
            var lastTableCell = worksheet.Cell(outputSheet.LastRowUsed().RowNumber(),
                outputSheet.LastColumnUsed().ColumnNumber());

            var rngData = outputSheet.Range(firstTableCell.Address, lastTableCell.Address);

            var row = worksheet.FirstRowUsed().RowNumber();
            var column = worksheet.FirstColumnUsed().ColumnNumber();

            newWorksheet.Cell(row, column).Value = rngData.CopyTo(newWorksheet.Cell(row, column)).FirstCell().Value;

            newWorkbook.SaveAs(Options.ResultFilePath);
        }
    }
}
