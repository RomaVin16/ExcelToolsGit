namespace ExcelTools.Abstraction
{
    public class ExcelOptionsBase
    {
        /// <summary>
        /// Имя нового файла 
        /// </summary>
        public string ResultFilePath { get; set; }

        /// <summary>
        /// Количество пропускаемных строк в файле
        /// </summary>
        public int SkipRows { get; set; } = 0;

        /// <summary>
        /// Номер листа для обработки
        /// </summary>
        public int SheetNumber { get; set; } = 1;

        /// <summary>
        /// Проверка корректности настроек 
        /// </summary>
        public virtual bool Validate()
        {
            return !string.IsNullOrEmpty(ResultFilePath);
        }
    }
}
