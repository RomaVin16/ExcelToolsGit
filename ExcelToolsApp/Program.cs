using ExcelTools.App;
using ExcelTools.ColumnSplitter;
using ExcelTools.Comparison;
using ExcelTools.Rotate;
using ExcelTools.Splitter;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            var comparison = new Comparison();

            var result = comparison.Process(new ComparisonOptions
            {
                SourceFilePath = "comparison1.xlsx",
                ModifiedFilePath = "comparison2.xlsx",
                KeysForRowsComparison = new[] { "E" },
                Id = new[] { "C" },
                ResultFilePath = "new_comparison.xlsx",
                SheetNumber = 1
            });

            //AppHelper.DeleteFolder();
        }
    }
}
