using ClosedXML.Excel;
using ExcelTools.Abstraction;
using ExcelTools.Exceptions;

namespace ExcelTools.ColumnSplitter
{
    public class ColumnSplitter: ExcelHandlerBase<ColumnSplitterOptions, ColumnSplitterResult>
    {
        public override ColumnSplitterResult Process(ColumnSplitterOptions options)
        {
            Options = options;

            if (!options.Validate())
            {
                return ErrorResult("Wrong options");
            }

            try
            {
                SplitColumn();

                return this.Result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        protected void SplitColumn()
        {
            using var workbook = new XLWorkbook(Options.FilePath);
            var worksheet = workbook.Worksheet(Options.SheetNumber);

            if (worksheet.Column(Options.ColumnName).IsEmpty())
            {
                throw new ExcelToolsException("Column is empty!");
            }

            string[] splitRow;

            var firstRowForProcessing = worksheet.Column(Options.ColumnName).FirstCellUsed().Address.RowNumber + Options.SkipHeaderRows;
            var lastRowForProcessing = worksheet.Column(Options.ColumnName).LastCellUsed().Address.RowNumber;

            var maxSplitLength = 0;
            for (var i = firstRowForProcessing; i <= lastRowForProcessing; i++)
            {
                splitRow = worksheet.Cell(i, Options.ColumnName).Value.ToString().Split(Options.SplitSymbols).ToArray();
                maxSplitLength = Math.Max(maxSplitLength, splitRow.Length);
            }

            Result.CreatedColumns = maxSplitLength;

            worksheet.Column(Options.ColumnName).InsertColumnsAfter(maxSplitLength);
            int column;

            for (var i = firstRowForProcessing; i <= lastRowForProcessing; i++)
            {
                splitRow = worksheet.Cell(i, Options.ColumnName).Value.ToString().Split(Options.SplitSymbols).ToArray();
                column = worksheet.Column(Options.ColumnName).ColumnNumber();

                foreach (var item in splitRow)
                {
                    worksheet.Cell(i, column + 1).Value = item;
                    column++;
                }

                Result.ProcessedRows++;
            }

            workbook.SaveAs(Options.ResultFilePath);
        }
    }
}
