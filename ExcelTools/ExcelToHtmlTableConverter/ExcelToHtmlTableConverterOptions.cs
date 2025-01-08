using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTools.Abstraction;

namespace ExcelTools.ExcelToHtmlTableConverter
{
    public class ExcelToHtmlTableConverterOptions: ExcelOptionsBase
    {
        /// <summary>
        /// Имя исходного файла 
        /// </summary>
        public string FilePath { get; set; }
    }
}
