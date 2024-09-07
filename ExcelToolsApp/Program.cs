using ExcelTools.App;
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
                ResultsCount = 20, 
                AddHeaderRows = 1,
                SplitMode = SplitterOptions.SplitType.SplitByRows

            });

            AppHelper.PrintProgramOperationStatistics(result);
        }
    }
}
