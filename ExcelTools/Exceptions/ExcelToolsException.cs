namespace ExcelTools.Exceptions
{
    public class ExcelToolsException: Exception
    {
        public ExcelToolsException(): base() { }

        public ExcelToolsException(string message) : base(message) { }

        public ExcelToolsException(string message, Exception innerException) : base(message, innerException) { }
    }
}
