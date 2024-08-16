using ExcelTools.App;
using ExcelTools.Cleaner;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main()
        {

            var cleanOptions = new CleanOptions { FilePath = "test.xlsx", ResultFilePath = "newTestEmptyFile.xlsx" };
            var cleanResult = new CleanResult();

            var cleaner = new Cleaner(cleanOptions, cleanResult);

            cleaner.Process();
            Helper.PrintProgramOperationStatistics(cleanResult);
        }
    }
}
