using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTools.Abstraction;

namespace ExcelTools.Comparison
{
    public class ComparisonResult: ExcelResultBase
    {
public int CountDeletedRows { get; set; }
public int CountAddedRows { get; set; }
public int CountChangedRows { get; set; }
    }
}
