using ExcelTools.App;
using ExcelTools.Merger;
using ExcelTools.Splitter;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main()
        {
            TestSplitter();
        }

        static void TestSplitter()
        {
            var splitter = new Splitter();

            var result = splitter.Process(new SplitterOptions()
            {
                FilePath = "split_input.xlsx",
                ResultFilePath = "new_file{0}.xlsx",
                SplitMode = SplitterOptions.SplitType.SplitByFiles,
                AddHeaderRows = 1,
                ResultsCount = 3
            });

            AppHelper.PrintProgramOperationStatistics(result);
        }
    }
}
