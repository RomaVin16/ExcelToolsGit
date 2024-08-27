using ExcelTools.Abstraction;

namespace ExcelTools.Merger
{
    public class MergerOptions: ExcelOptionsBase
    {
        public string[] MergeFilePaths { get; set; }
    }
}
