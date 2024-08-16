namespace ExcelTools.Abstraction
{
    public class ExcelResultBase
    {
        public enum ResultCode
        {
            Success,
            Error
        }

        public ResultCode Code { get; set; }
        public string ErrorMessage { get; set; }
    }
}
