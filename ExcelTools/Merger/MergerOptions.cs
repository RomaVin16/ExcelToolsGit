using ExcelTools.Abstraction;

namespace ExcelTools.Merger
{
    public class MergerOptions: ExcelOptionsBase
    {
        /// <summary>
        /// Имена файлов для слияния 
        /// </summary>
        public string[] MergeFilePaths { get; set; }

        /// <summary>
        /// Тип копируемых данных
        /// </summary>
        public enum MergeType
        {
            Table,
            Sheets
        }

        public MergeType MergeMode { get; set; } = MergeType.Table;

        /// <summary>
        /// Проверка корректности настроек 
        /// </summary>
        public override bool Validate()
        {
            return !string.IsNullOrEmpty(ResultFilePath) && MergeFilePaths.Length > 1;
        }
    }
}
