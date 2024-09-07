using ExcelTools.Abstraction;

namespace ExcelTools.ColumnSplitter
{
    public class ColumnSplitterOptions: ExcelOptionsBase
    {
        public string FilePath { get; set; }
        public string SplitSymbols { get; set; }
        public string ColumnName { get; set; } = "A";
        public int SkipHeaderRows { get; set; }
    }
}
