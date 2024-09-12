using ExcelTools.Abstraction;

namespace ExcelTools.ColumnSplitter
{
    public class ColumnSplitterOptions: ExcelOptionsBase
    {
        /// <summary>
        /// Имя исходного файла 
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// Символ, по которому требуется разделить строки на столбцы 
        /// </summary>
        public string SplitSymbols { get; set; }

        /// <summary>
        /// Имя столбца, в котором требуется произвести разделение 
        /// </summary>
        public string ColumnName { get; set; } = "A";

        /// <summary>
        /// Количество строк (заголовков), которые нужно пропустить 
        /// </summary>
        public int SkipHeaderRows { get; set; }
    }
}
