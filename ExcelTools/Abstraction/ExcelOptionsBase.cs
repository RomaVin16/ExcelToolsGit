namespace ExcelTools.Abstraction
{
    public class ExcelOptionsBase
    {
        public string ResultFilePath { get; set; }
        public int SkipRows { get; set; } = 0;

        public virtual bool Validate()
        {
            return !string.IsNullOrEmpty(ResultFilePath);
        }
    }
}
