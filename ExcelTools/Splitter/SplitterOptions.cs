using ExcelTools.Abstraction;

namespace ExcelTools.Splitter
{
    public class SplitterOptions: ExcelOptionsBase
    {
        public string FilePath { get; set; }
        public int ResultsCount { get; set; }
        public int AddHeaderRows { get; set; }

        public enum SplitType
        {
            SplitByRows, 
            SplitByFiles
        }

        public SplitType SplitMode { get; set; } = SplitType.SplitByRows;

        public override bool Validate()
        {
            return !string.IsNullOrEmpty(FilePath);
        }
    }
}
