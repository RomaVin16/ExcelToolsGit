namespace ExcelTools
{
    /// <summary>
    /// статистика работы программы
    /// </summary>
    public class CleanResult
    {
        public enum ResultCode
        {
            Success,
            Error
        }

        public ResultCode Code { get; set; }
        public string ErrorMessage { get; set; }
        public int RowsProcessed { get; set; }
        public int RowsRemoved { get; set; }
    }
}
