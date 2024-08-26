using ExcelTools.App;
using ExcelTools.DuplicateRemover;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main()
        {
            TestCleaner();
            Console.ReadLine();
        }

        static void TestCleaner()
        {
            var remover = new DuplicateRemover();

            var result = remover.Process(new DuplicateRemoverOptions() { 
                FilePath = "test.xlsx", 
                ResultFilePath = "test-result.xlsx",
            });

            AppHelper.PrintProgramOperationStatistics(result);
        }
    }
}
