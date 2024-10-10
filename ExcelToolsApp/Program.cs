using ExcelTools.App;
using ExcelTools.ColumnSplitter;
using ExcelTools.Rotate;
using ExcelTools.Splitter;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main(string[] args)
        {

            var rotater = new Rotater();

            var Result = rotater.Process(new RotaterOptions
            {
                FilePath = "rotate.xlsx",
                ResultFilePath = "new_rotate.xlsx",
                SheetNumber = 1,
                SkipRows = 0
            });

            //AppHelper.DeleteFolder();
        }
    }
}
