using ExcelTools.Abstraction;

namespace ExcelTools.ColumnSplitter
{
    public class ColumnSplitterResult: ExcelResultBase
    {
        public int ProcessedRows { get; set; }
        public int CreatedColumns { get; set; }
    }
}
