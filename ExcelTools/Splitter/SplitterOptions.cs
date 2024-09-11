using ExcelTools.Abstraction;

namespace ExcelTools.Splitter
{
    public class SplitterOptions: ExcelOptionsBase
    {
        /// <summary>
        /// Имя исходного файла 
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// Количество строк/файлов в новом созданном файле
        /// </summary>
        public int ResultsCount { get; set; }

        /// <summary>
        /// Количество строк, которые требуется использовать как заголовки 
        /// </summary>
        public int AddHeaderRows { get; set; }

        /// <summary>
        /// Критерий, по которому происходит разделение данных 
        /// </summary>
        public enum SplitType
        {
            SplitByRows, 
            SplitByFiles
        }

        public SplitType SplitMode { get; set; } = SplitType.SplitByRows;

        /// <summary>
        /// Проверка корректности настроек 
        /// </summary>
        public override bool Validate()
        {
            return !string.IsNullOrEmpty(FilePath);
        }
    }
}
