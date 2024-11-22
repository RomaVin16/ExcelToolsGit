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
            using var sourceWorkbook = new XLWorkbook(Options.SourceFilePath);
var sourceWorksheet = sourceWorkbook.Worksheet(Options.SheetNumber);

using var modifiedWorkbook = new XLWorkbook(Options.ModifiedFilePath);
var modifiedWorksheet = modifiedWorkbook.Worksheet(Options.SheetNumber);

using var newWorkbook = new XLWorkbook();
modifiedWorksheet.CopyTo(newWorkbook, modifiedWorksheet.Name);
var newWorksheet = newWorkbook.Worksheet(1);

newWorkbook.AddWorksheet("Changed");
newWorkbook.AddWorksheet("Deleted");
newWorkbook.AddWorksheet("Added");

ComparisonHelper.InsertHeadersInList(newWorksheet, newWorkbook, Options.HeaderRows);

newWorkbook.SaveAs(Options.ResultFilePath);

if (Options.Id != null)
{
    var comparisonWithKey= new ComparisonWithKey(Options, Result);
    comparisonWithKey.CompareByKey(newWorkbook, sourceWorksheet, newWorksheet);
}
else
{
    var comparisonDefault = new ComparisonDefault(Options, Result);
    comparisonDefault.CompareWithoutKey(newWorkbook, sourceWorksheet, newWorksheet);
}
        }
    }
}

