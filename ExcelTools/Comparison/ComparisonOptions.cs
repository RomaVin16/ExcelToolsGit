using ExcelTools.Abstraction;

namespace ExcelTools.Comparison
{
    public class ComparisonOptions: ExcelOptionsBase
    {
        /// <summary>
        /// Имя исходного файла 
        /// </summary>
        public string SourceFilePath { get; set; }

        /// <summary>
        /// Имя исходного файла 
        /// </summary>
        public string ModifiedFilePath { get; set; }

        public string[]? Id { get; set; }

        /// <summary>
        /// Номера, которые требуется использовать как заголовки 
        /// </summary>
        public int[] HeaderRows { get; set; }
    }
}
