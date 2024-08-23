using ExcelTools.Abstraction;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTools.DuplicateRemover
{
    public class DuplicateRemoverOptions : ExcelOptionsBase
    {
        public string[] KeysForRowsComparison { get; set; }
    }
}