using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ExcelTools.Abstraction;
using ExcelTools.Splitter;

namespace ExcelTools.ColumnSplitter
{
    public class ColumnSplitter: ExcelHandlerBase<ColumnSplitterOptions, ColumnSplitterResult>
    {
ColumnSplitterOptions options = new ColumnSplitterOptions();
private ColumnSplitterResult result = new ColumnSplitterResult();

        public override ColumnSplitterResult Process(ColumnSplitterOptions options)
        {
            this.options = options;

            if (!options.Validate())
            {
                return ErrorResult("Wrong options");
            }

            try
            {
                
                SplitColumn();

                return this.result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        protected void SplitColumn()
        {
            var workbook = new XLWorkbook(options.FilePath);
            var worksheet = workbook.Worksheet(options.SheetNumber);

            var columnName = options.ColumnName;
            string[] splitRow;

            var maxSplitLength = 0;
            for (var i = worksheet.Column(columnName).FirstCellUsed().Address.RowNumber + options.SkipHeaderRows; i <= worksheet.Column(columnName).LastCellUsed().Address.RowNumber; i++)
            {
                splitRow = worksheet.Cell(i, columnName).Value.ToString().Split(options.SplitSymbols).ToArray();
                maxSplitLength = Math.Max(maxSplitLength, splitRow.Length);
            }

            result.CreatedColumns = maxSplitLength;

            worksheet.Column(columnName).InsertColumnsAfter(maxSplitLength);

            for (var i = worksheet.Column(columnName).FirstCellUsed().Address.RowNumber + options.SkipHeaderRows; i <= worksheet.Column(columnName).LastCellUsed().Address.RowNumber; i++)
            {
                splitRow = worksheet.Cell(i, columnName).Value.ToString().Split(options.SplitSymbols).ToArray();
                var column = worksheet.Column(columnName).ColumnNumber();

                foreach (var item in splitRow)
                {
                    worksheet.Cell(i,column + 1).Value = item;
                    column++;
                }

                result.ProcessedRows++;
            }

            workbook.SaveAs(options.ResultFilePath);
        }
    }
}
