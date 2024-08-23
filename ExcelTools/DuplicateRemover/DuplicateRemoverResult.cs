using ExcelTools.Abstraction;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemoverResult : ExcelResultBase
    {
        public int RowsProcessed { get; set; }
        public int RowsRemoved { get; set; }
    }
}
