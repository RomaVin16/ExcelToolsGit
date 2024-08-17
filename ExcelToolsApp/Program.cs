using ExcelTools.App;
using ExcelTools.Cleaner;

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
            var cleaner = new Cleaner();

            var result = cleaner.Process(new CleanOptions { 
                FilePath = "test.xlsx", 
                ResultFilePath = "test-result.xlsx" });

            AppHelper.PrintProgramOperationStatistics(result);

        }
    }
}
