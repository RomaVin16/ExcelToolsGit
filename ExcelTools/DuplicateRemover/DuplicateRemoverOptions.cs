using ExcelTools.Abstraction;

namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemoverOptions : ExcelOptionsBase
    {
        public string FilePath { get; set; }
        public string[] KeysForRowsComparison { get; set; }
    }
}