using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APILib.APIOptions
{
    public class DuplicateRemoverAPIOptions: APIOptionsBase
    {
        public Guid FileId { get; set; }

        /// <summary>
        /// Ключ для поиска дубликатов 
        /// </summary>
        [DefaultValue(null)]
        public string[]? KeysForRowsComparison { get; set; } = null;
    }
}
