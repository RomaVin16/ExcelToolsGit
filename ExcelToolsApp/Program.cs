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
                SourceFilePath = "test7.xlsx",
                ModifiedFilePath = "test7-changed.xlsx",
                Id = new[] { "A" },
                ResultFilePath = "new_comparison.xlsx",
                SheetNumber = 1,
                HeaderRows = new[] { 2 }
            });
        }
    }
}
