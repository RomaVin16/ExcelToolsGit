using ClosedXML.Excel;
using ExcelTools.Abstraction;

namespace ExcelTools.Splitter
{
    public class Splitter: ExcelHandlerBase<SplitterOptions, SplitterResult>
    {

        public override SplitterResult Process(SplitterOptions options)
        {
            Options = options;

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
                else if (options.SplitMode == SplitterOptions.SplitType.SplitByRows)
                {
                    SplitFileByNumberOfRows();
                }
                else
                {
                    return ErrorResult("This split method is not supported!");
                }

                return Result;
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
            using var workbook = new XLWorkbook(Options.FilePath);
        var worksheet = workbook.Worksheet(Options.SheetNumber);

        var firstRowNumber = worksheet.FirstRowUsed().RowNumber();
        var firstColumnNumber = worksheet.FirstColumnUsed().ColumnNumber();

        var firstRowNumberInRange = firstRowNumber + Options.AddHeaderRows;

            var numberOfRowInResultFiles = Math.Ceiling((double)(worksheet.RowsUsed().Count() - Options.AddHeaderRows) / Options.ResultsCount);

        var lastRowNumber = firstRowNumber - 1 + Options.AddHeaderRows + numberOfRowInResultFiles;
        var lastColumnNumber = worksheet.LastColumnUsed().ColumnNumber();

        var headerRange = worksheet.Range(firstRowNumber, firstColumnNumber, (int)lastRowNumber, lastColumnNumber);

        CopyDataInSheet(worksheet, firstRowNumberInRange, firstColumnNumber, lastRowNumber, lastColumnNumber, headerRange, firstRowNumber, numberOfRowInResultFiles);
        }

    /// <summary>
    /// Разделить таблицу по количеству строк
    /// </summary>
    protected void SplitFileByNumberOfRows()
        {
            using var workbook = new XLWorkbook(Options.FilePath);
            var worksheet = workbook.Worksheet(Options.SheetNumber);

            var firstRowNumber = worksheet.FirstRowUsed().RowNumber();
            var firstColumnNumber = worksheet.FirstColumnUsed().ColumnNumber();

            var lastRowNumber =  firstRowNumber - 1 + Options.ResultsCount + Options.AddHeaderRows;
            var lastColumnNumber = worksheet.LastColumnUsed().ColumnNumber();

            var firstRowNumberInRange = firstRowNumber + Options.AddHeaderRows;

            var headerRange = worksheet.Range(firstRowNumber, firstColumnNumber, lastRowNumber, lastColumnNumber);

            var totalFiles = (int)Math.Ceiling((double)(worksheet.RowsUsed().Count() - Options.AddHeaderRows) / Options.ResultsCount);

            CopyDataInSheet(worksheet, firstRowNumberInRange, firstColumnNumber, lastRowNumber, lastColumnNumber,
                headerRange, firstRowNumber, totalFiles);
        }

    /// <summary>
    /// Копирование данных 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="firstRowNumberInRange"></param>
    /// <param name="firstColumnNumber"></param>
    /// <param name="lastRowNumber"></param>
    /// <param name="lastColumnNumber"></param>
    /// <param name="headerRange"></param>
    /// <param name="firstRowNumber"></param>
    /// <param name="numberOfRowInResultFiles"></param>
    protected void CopyDataInSheet(IXLWorksheet worksheet, int firstRowNumberInRange, int firstColumnNumber, double lastRowNumber, int lastColumnNumber, IXLRange headerRange, int firstRowNumber, double numberOfRowInResultFiles)
    {
        for (var i = 1; i <= Options.ResultsCount; i++)
        {
            using var newWorkbook = new XLWorkbook();
            var newWorksheet = newWorkbook.AddWorksheet("Sheet1");

            var rngData = worksheet.Range(firstRowNumberInRange, firstColumnNumber, (int)lastRowNumber, lastColumnNumber);

            headerRange.CopyTo(newWorksheet.Cell(firstRowNumber, firstColumnNumber));
            rngData.CopyTo(newWorksheet.Cell(firstRowNumber + Options.AddHeaderRows, firstColumnNumber));

            newWorkbook.SaveAs(Options.IndividualPathToEachFile == null
                ? string.Format(Options.ResultFilePath, i)
                : Options.IndividualPathToEachFile(i));


            Result.NumberOfResultFiles++;

            firstRowNumberInRange += (int)numberOfRowInResultFiles;
            lastRowNumber += (int)numberOfRowInResultFiles;
        }
        }

    /// <summary>
    /// Копирование данных 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="firstRowNumberInRange"></param>
    /// <param name="firstColumnNumber"></param>
    /// <param name="lastRowNumber"></param>
    /// <param name="lastColumnNumber"></param>
    /// <param name="headerRange"></param>
    /// <param name="firstRowNumber"></param>
    /// <param name="totalFiles"></param>
    protected void CopyDataInSheet(IXLWorksheet worksheet, int firstRowNumberInRange, int firstColumnNumber, int lastRowNumber, int lastColumnNumber, IXLRange headerRange, int firstRowNumber, int totalFiles)
    {
        for (var i = 1; i <= totalFiles; i++)
        {
            using var newWorkbook = new XLWorkbook();
            var newWorksheet = newWorkbook.AddWorksheet("Sheet1");

            var rngData = worksheet.Range(firstRowNumberInRange, firstColumnNumber, lastRowNumber, lastColumnNumber);

            headerRange.CopyTo(newWorksheet.Cell(firstRowNumber, firstColumnNumber));
            rngData.CopyTo(newWorksheet.Cell(firstRowNumber + Options.AddHeaderRows, firstColumnNumber));

            newWorkbook.SaveAs(string.Format(Options.ResultFilePath, i));
            Result.NumberOfResultFiles++;

            firstRowNumberInRange += Options.ResultsCount;
            lastRowNumber += Options.ResultsCount;
        }
        }
    }
}


