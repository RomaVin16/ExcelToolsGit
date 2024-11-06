using ExcelTools.Comparison;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            var comparison = new Comparison();

            var result = comparison.Process(new ComparisonOptions
            {
                SourceFilePath = "test6.xlsx",
                ModifiedFilePath = "test6-changed.xlsx",
                Id = new[] { "A" },
                ResultFilePath = "new_comparison.xlsx",
                SheetNumber = 1,
                HeaderRows = new[] { 2 }
            });
        }
    }
}
