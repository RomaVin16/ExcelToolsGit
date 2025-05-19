using System.ComponentModel;

namespace APILib.APIOptions
{
    public class ColumnSplitterAPIOptions: APIOptionsBase
    {
        public Guid FileId { get; set; }

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
        [DefaultValue(1)]
        public int AddHeaderRows { get; set; } = 1;
    }
}
