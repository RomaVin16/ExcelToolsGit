using System.ComponentModel;

namespace APILib.APIOptions
{
    public class APIOptionsBase
    {
        /// <summary>
        /// Количество пропускаемных строк в файле
        /// </summary>
        [DefaultValue(0)]
        public int SkipRows { get; set; } = 0;

        /// <summary>
        /// Номер листа для обработки
        /// </summary>
        [DefaultValue(1)]
        public int SheetNumber { get; set; } = 1;
    }
}
