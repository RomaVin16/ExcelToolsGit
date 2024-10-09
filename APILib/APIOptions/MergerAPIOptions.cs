using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APILib.APIOptions
{
    public class MergerAPIOptions : APIOptionsBase
    {
        /// <summary>
        /// Имена файлов для слияния 
        /// </summary>
        public Guid[] MergeFilePaths { get; set; }

        /// <summary>
        /// Тип копируемых данных
        /// </summary>
        public enum MergeType
        {
            Table,
            Sheets
        }

        [DefaultValue(MergeType.Table)]
        public MergeType MergeMode { get; set; } = MergeType.Table;
    }
}
