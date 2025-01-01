using ExcelTools.Converter;

namespace ExcelToolsApp
{
    public class Program
    {
        static void Main()
        {
            var converter = new JsonToExcelConverter();

            var result = converter.Process(new JsonToExcelConverterOptions
            {
                FilePath = "testJson4.json",
                ResultFilePath = "new_converter.xlsx"
            });
        }
    }
}

