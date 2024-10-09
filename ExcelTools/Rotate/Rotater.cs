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
            var worksheet = workbook.Worksheet(Options.SheetNumber);

            var worksheet1 = workbook.AddWorksheet("output");

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
                            var style = worksheet.Cell(mergedRange.FirstRow().RowNumber(), mergedRange.FirstColumn().ColumnNumber()).Style;

                            var newMergedRange = worksheet1.Range(worksheet1.Cell(j, i), worksheet1.Cell(j + mergedRange.LastColumn().ColumnNumber() - mergedRange.FirstColumn().ColumnNumber(), i + mergedRange.LastRow().RowNumber() - mergedRange.FirstRow().RowNumber()));
                            newMergedRange.Merge();

                            worksheet1.Cell(j, i).Value = value;
                            worksheet1.Cell(j, i).Style = style;

                            j += mergedRange.LastColumn().ColumnNumber() - mergedRange.FirstColumn().ColumnNumber();
                        }
                    }
                    else
                    {
                        var value = worksheet.Cell(i, j).Value;
                        var style = worksheet.Cell(i, j).Style;
                        worksheet1.Cell(j, i).Value = value;
                        worksheet1.Cell(j, i).Style = style;
                    }
                }
            }



            workbook.SaveAs(Options.ResultFilePath);
        }
    }
}
