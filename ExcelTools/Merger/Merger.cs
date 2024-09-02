using ClosedXML.Excel;
using ExcelTools.Abstraction;
using ExcelTools.Exceptions;

namespace ExcelTools.Merger
{
    public class Merger : ExcelHandlerBase<MergerOptions, MergerResult>
    {
        private MergerOptions options = new MergerOptions();
        private MergerResult result = new MergerResult();

        public override MergerResult Process(MergerOptions options)
        {
            this.options = options;

            if (!options.Validate())
            {
                return ErrorResult("Wrong options");
            }

            try
            {
                MergeDataFromFiles(options.MergeFilePaths);

                return this.result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        protected void MergeDataFromFiles(string[] mergeFilePath)
        {
            using var workbook = new XLWorkbook(mergeFilePath[0]);

            if (!options.Validate())
            {
                throw new ExcelToolsException("Wrong options");
            }

            var mainWorksheet = workbook.Worksheet(1);

            result.NumberOfMergedFiles++;

            for (var i = 1; i < mergeFilePath.Length; i++)
            {
                using var mergeWorkbook = new XLWorkbook(mergeFilePath[i]);
                var worksheet = mergeWorkbook.Worksheet(1);

                if (worksheet.IsEmpty())
                {
                    continue;
                }

                if (options.MergeMode == MergerOptions.MergeType.Table)
                {
                    CopyData(mainWorksheet, worksheet);
                    result.NumberOfMergedFiles++;
                }
                else
                {
                    CopySheetsInMainFile(workbook, mergeWorkbook, i);
                    result.NumberOfMergedFiles++;
                }
            }

            workbook.SaveAs(options.ResultFilePath);
        }

        /// <summary>
        /// Копирование диапазона из одного файла в другой
        /// </summary>
        /// <param name="mainWorksheet"></param>
        /// <param name="worksheet"></param>
        protected void CopyData(IXLWorksheet mainWorksheet, IXLWorksheet worksheet)
        {
            var firstTableCell = worksheet.Cell(worksheet.FirstRowUsed().RowNumber() + options.SkipRows,
                worksheet.FirstColumnUsed().ColumnNumber());
            var lastTableCell = worksheet.Cell(worksheet.LastRowUsed().RowNumber(),
                worksheet.LastColumnUsed().ColumnNumber());

            var rngData = worksheet.Range(firstTableCell.Address, lastTableCell.Address);

            var row = mainWorksheet.LastRowUsed().RowNumber() + 1;
            var column = mainWorksheet.FirstColumnUsed().ColumnNumber();

            mainWorksheet.Cell(row, column).Value = rngData.CopyTo(mainWorksheet.Cell(row, column)).FirstCell().Value;
        }

        /// <summary>
        /// Копирование листов между файлами
        /// </summary>
        /// <param name="mainWorkbook"></param>
        /// <param name="currentWorkbook"></param>
        /// <param name="fileNumber"></param>
        protected void CopySheetsInMainFile(IXLWorkbook mainWorkbook, IXLWorkbook currentWorkbook, int fileNumber)
        {
            for (var i = 1; i <= currentWorkbook.Worksheets.Count; i++)
            {
                currentWorkbook.Worksheet(i).CopyTo((XLWorkbook)mainWorkbook, $"File{fileNumber + 1}_Sheet{i}");
            }
        }
    }
}
