using ExcelTools.Abstraction;

namespace ExcelTools.Cleaner
{
    public class CleanResult : ExcelResultBase
    {
        public int RowsProcessed { get; set; }
        public int RowsRemoved { get; set; }
    }
}
