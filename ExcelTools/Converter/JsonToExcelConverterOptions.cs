using ExcelTools.Abstraction;

namespace ExcelTools.Converter
{
public class JsonToExcelConverterOptions: ExcelOptionsBase
    {
        /// <summary>
        /// Имя исходного файла 
        /// </summary>
        public string FilePath { get; set; }
    }
}
