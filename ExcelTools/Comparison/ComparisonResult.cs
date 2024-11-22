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
        /// <summary>
        /// Количество удаленных строк 
        /// </summary>
public int CountDeletedRows { get; set; }

        /// <summary>
        /// Количество добавленных строк 
        /// </summary>
        public int CountAddedRows { get; set; }

        /// <summary>
        /// Количество измененных строк 
        /// </summary>
        public int CountChangedRows { get; set; }

        /// <summary>
        /// Количество удаленных столбцов  
        /// </summary>
        public int CountDeletedColumns { get; set; }

        /// <summary>
        /// Количество доавлнных столбцов 
        /// </summary>
        public int CountAddedColumns { get; set; }
    }
}
