using ExcelTools.App;
using ExcelTools.Merger;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main()
        {
            TestCleaner();
        }

        static void TestCleaner()
        {
            var merger = new Merger();

            var result = merger.Process(new MergerOptions
            { 
                MergeFilePaths = new []{ "merge_1.xlsx","merge_2.xlsx","merge_3.xlsx" },
                ResultFilePath = "new_merge_result.xlsx",
                SkipRows = 1, 
                MergeMode = MergerOptions.MergeType.Table
            });

            AppHelper.PrintProgramOperationStatistics(result);
        }
    }
}
