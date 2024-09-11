using ExcelTools.Abstraction;

namespace ExcelTools.Cleaner
{
    public class CleanOptions: ExcelOptionsBase
    {
        /// <summary>
        /// Имя исходного файла 
        /// </summary>
        public string FilePath { get; set; }
    }
}
