using ExcelTools.Abstraction;

namespace ExcelTools.Merger
{
    public class MergerOptions: ExcelOptionsBase
    {
        public string[] MergeFilePaths { get; set; }

        public enum MergeType
        {
            Table,
            Sheets
        }

        public MergeType MergeMode { get; set; } = MergeType.Table;

        public override bool Validate()
        {
            return !string.IsNullOrEmpty(ResultFilePath) && MergeFilePaths.Length > 1;
        }
    }
}
