namespace ExcelTools.Abstraction
{
    public abstract class ExcelHandlerBase<TOptions, TResult> 
        where TOptions: ExcelOptionsBase 
        where TResult : ExcelResultBase, new()
    {
        public abstract TResult Process(TOptions options);

        public TResult ErrorResult(string message)
        {
            return new TResult() { Code = ResultCode.Error, ErrorMessage = message };
        }
       
    }
}
