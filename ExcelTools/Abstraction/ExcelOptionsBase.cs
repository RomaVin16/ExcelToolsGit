using DocumentFormat.OpenXml.Office.CoverPageProps;

namespace ExcelTools.Abstraction
{
    public class ExcelOptionsBase
    {
        public string FilePath { get; set; }
        public string ResultFilePath { get; set; }
        public int SkipRows { get; set; } = 0;

        public virtual bool Validate()
        {
            return !string.IsNullOrEmpty(FilePath) && !string.IsNullOrEmpty(ResultFilePath);

        }
    }
}
