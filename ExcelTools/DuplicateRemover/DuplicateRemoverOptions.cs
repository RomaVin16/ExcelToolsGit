using ExcelTools.Abstraction;

namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemoverOptions : ExcelOptionsBase
    {
        public int SkipRows { get; set; } = 0;

        public override bool Validate()
        {
            base.Validate();
            if (SkipRows is < 0 or > 65536)
            {
                return false;
            }

            return true;
        }
    }
}