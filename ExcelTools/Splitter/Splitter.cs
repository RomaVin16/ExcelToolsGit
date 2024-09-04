using ClosedXML.Excel;
using ExcelTools.Abstraction;

namespace ExcelTools.Splitter
{
    public class Splitter: ExcelHandlerBase<SplitterOptions, SplitterResult>
    {
        private SplitterOptions options = new SplitterOptions();
        private SplitterResult result = new SplitterResult();

        public override SplitterResult Process(SplitterOptions options)
        {
            this.options = options;

            if (!options.Validate())
            {
                return ErrorResult("Wrong options");
            }

            try
            {
                if (options.SplitMode == SplitterOptions.SplitType.SplitByFiles)
                {
                    SplitFileByNumberOfFiles();
                }
                else
                {
                    SplitFileByNumberOfRows();
                }

                return this.result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        /// <summary>
        /// Разделить таблицу по количеству файлов
        /// </summary>
        protected void SplitFileByNumberOfFiles()
        {
            using var workbook = new XLWorkbook(options.FilePath);
        var worksheet = workbook.Worksheet(1);

        var firstRowNumber = worksheet.FirstRowUsed().RowNumber();
        var firstColumnNumber = worksheet.FirstColumnUsed().ColumnNumber();

        var firstRowNumberInRange = firstRowNumber + options.AddHeaderRows;

            var numberOfRowInResultFiles = Math.Ceiling((double)(worksheet.RowsUsed().Count() - options.AddHeaderRows) / options.ResultsCount);

        var lastRowNumber = firstRowNumber - 1 + options.AddHeaderRows + numberOfRowInResultFiles;
        var lastColumnNumber = worksheet.LastColumnUsed().ColumnNumber();

        var headerRange = worksheet.Range(firstRowNumber, firstColumnNumber, (int)lastRowNumber, lastColumnNumber);

        for (var i = 1; i <= options.ResultsCount; i++)
        {
            var newWorkbook = new XLWorkbook();
            var newWorksheet = newWorkbook.AddWorksheet("Sheet1");

            var rngData = worksheet.Range(firstRowNumberInRange, firstColumnNumber, (int)lastRowNumber, lastColumnNumber);

                headerRange.CopyTo(newWorksheet.Cell(firstRowNumber, firstColumnNumber));
                rngData.CopyTo(newWorksheet.Cell(firstRowNumber + options.AddHeaderRows, firstColumnNumber));

            newWorkbook.SaveAs(string.Format(options.ResultFilePath, i));
            result.NumberOfResultFiles++;

            firstRowNumberInRange += (int)numberOfRowInResultFiles;
            lastRowNumber += (int)numberOfRowInResultFiles;
        }
        }

    /// <summary>
    /// Разделить таблицу по количеству строк
    /// </summary>
    protected void SplitFileByNumberOfRows()
        {
            using var workbook = new XLWorkbook(options.FilePath);
            var worksheet = workbook.Worksheet(1);

            var firstRowNumber = worksheet.FirstRowUsed().RowNumber();
            var firstColumnNumber = worksheet.FirstColumnUsed().ColumnNumber();

            var lastRowNumber =  firstRowNumber - 1 + options.ResultsCount + options.AddHeaderRows;
            var lastColumnNumber = worksheet.LastColumnUsed().ColumnNumber();

            var firstRowNumberInRange = firstRowNumber + options.AddHeaderRows;

            var headerAddress = worksheet.Range(firstRowNumber, firstColumnNumber, lastRowNumber, lastColumnNumber);

            var totalFiles = (int)Math.Ceiling((double)(worksheet.RowsUsed().Count() - options.AddHeaderRows) / options.ResultsCount);

            for (var i = 1; i <= totalFiles; i++)
            {
                var newWorkbook = new XLWorkbook();
                var newWorksheet = newWorkbook.AddWorksheet("Sheet1");

                var rngData = worksheet.Range(firstRowNumberInRange, firstColumnNumber, lastRowNumber, lastColumnNumber);

                    headerAddress.CopyTo(newWorksheet.Cell(firstRowNumber, firstColumnNumber));
                    rngData.CopyTo(newWorksheet.Cell(firstRowNumber + options.AddHeaderRows, firstColumnNumber));

                newWorkbook.SaveAs(string.Format(options.ResultFilePath, i));
                result.NumberOfResultFiles++;

                firstRowNumberInRange += options.ResultsCount;
                lastRowNumber += options.ResultsCount;
            }
        }
    }
}
