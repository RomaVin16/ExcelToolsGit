using ExcelTools.Abstraction;

namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemoverResult : ExcelResultBase
    {
        public int RowsProcessed { get; set; }
        public int RowsRemoved { get; set; }
    }
}
