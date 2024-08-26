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
                FilePath = "test3.xlsx", 
                ResultFilePath = "test-result.xlsx", 
                KeysForRowsComparison = new []{ "B","C" }
            });

            AppHelper.PrintProgramOperationStatistics(result);
        }
    }
}
