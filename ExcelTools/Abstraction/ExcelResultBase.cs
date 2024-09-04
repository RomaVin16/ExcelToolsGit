namespace ExcelTools.Abstraction
{
    public enum ResultCode
    {
        Success,
        Error
    }
    public class ExcelResultBase
    {

        public ResultCode Code { get; set; }
        public string ErrorMessage { get; set; }

        public static ExcelResultBase ErrorResult(string message)
        {
            return new ExcelResultBase { Code = ResultCode.Error, ErrorMessage = message };
        }
    }
}
