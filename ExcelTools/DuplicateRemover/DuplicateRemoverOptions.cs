using ExcelTools.Abstraction;

namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemoverOptions : ExcelOptionsBase
    {
        public string[] KeysForRowsComparison { get; set; }
    }
}