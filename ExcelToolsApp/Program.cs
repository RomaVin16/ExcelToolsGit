using ExcelTools.App;
using ExcelTools.Cleaner;
using ExcelTools.DuplicateRemover;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main()
        {
            TestDuplicateRemover();
            Console.ReadLine();
        }

        static void TestDuplicateRemover()
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
