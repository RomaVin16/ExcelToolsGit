using ExcelTools.Abstraction;

namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemoverOptions : ExcelOptionsBase
    {
        /// <summary>
        /// Имя исходного файла 
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// Ключ для поиска дубликатов 
        /// </summary>
        public string[]? KeysForRowsComparison { get; set; }
    }
}