using ExcelTools;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main()
        {
            var cleaner = new Cleaner();

            var cleanerOptions = new CleanOptions();

            cleanerOptions.FilePath = "test";
            cleanerOptions.ResultFilePath = "newTestEmptyFile";
            cleanerOptions.CleanWhiteSpaceRows = false;

            CleanResult result = new CleanResult();

            try
            {
                result = cleaner.DeleteEmptyStringInWorkbook(cleanerOptions);
            }
            catch (Exception ex)
            {
                result.Code = (CleanResult.ResultCode)1;
                result.ErrorMessage = ex.Message;
            }

            ExcelHelper.PrintProgramOperationStatistics(result);
        }
    }
}
