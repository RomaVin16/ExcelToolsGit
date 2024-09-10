namespace ExcelTools.Abstraction
{
    public abstract class ExcelHandlerBase<TOptions, TResult> 
        where TOptions: ExcelOptionsBase, new() 
        where TResult : ExcelResultBase, new()
    {
        public abstract TResult Process(TOptions options);

        protected TOptions Options { get; set; }
        protected TResult Result { get; private set; }

        protected ExcelHandlerBase()
        {
            Options = new TOptions();
            Result = new TResult();
        }

        public TResult ErrorResult(string message)
        {
            return new TResult() { Code = ResultCode.Error, ErrorMessage = message };
        }

    }
}
