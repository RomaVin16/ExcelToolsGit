using ExcelTools.App;
using ExcelTools.ColumnSplitter;
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
            var splitter = new ColumnSplitter();

            var result = splitter.Process(new ColumnSplitterOptions()
            {
                FilePath = "column_split_input.xlsx",
                ResultFilePath = "new_file.xlsx", 
                ColumnName = "D", 
                SheetNumber = 1, 
                SkipHeaderRows = 1, 
                SplitSymbols = " "
            });

            AppHelper.PrintProgramOperationStatistics(result);
        }
    }
}
