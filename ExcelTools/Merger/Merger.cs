using System.Runtime.CompilerServices;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using ExcelTools.Abstraction;

namespace ExcelTools.Merger
{
    public class Merger: ExcelHandlerBase<MergerOptions, MergerResult>
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
    var firstWorksheet = workbook.Worksheet(1);

            for (var i = 1; i < mergeFilePath.Length; i++)
    {
        using var mergeWorkbook = new XLWorkbook(mergeFilePath[i]);
        var worksheet = mergeWorkbook.Worksheet(1);

        if (worksheet.IsEmpty())
        {
            continue;
        }

        InsertDataIntoSheet(firstWorksheet, worksheet);
    }

    workbook.SaveAs(options.ResultFilePath);
}

protected void InsertDataIntoSheet(IXLWorksheet firstWorksheet, IXLWorksheet worksheet)
{
    var targetRow = firstWorksheet.LastRowUsed().RowNumber() + 1;
    var targetColumn = firstWorksheet.FirstColumnUsed().ColumnNumber();

            var sourceRow = worksheet.FirstRowUsed().RowNumber() + options.SkipRows;

            for (var i = targetRow; i <= worksheet.RowsUsed().Count() - options.SkipRows + targetRow; i++)
    {
        var firstColumnInCurrentFile = worksheet.FirstColumnUsed().ColumnNumber();

        for (var j = targetColumn; j <= worksheet.ColumnsUsed().Count() + 1; j++)
        {
            firstWorksheet.Cell(i, j).Value = worksheet.Cell(sourceRow, firstColumnInCurrentFile).Value;
            firstColumnInCurrentFile++;
        }

        sourceRow++;
    }
}
    }
    }
