using ClosedXML.Excel;
using ExcelTools.Abstraction;
using ExcelTools.Exceptions;
using ExcelTools.Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools.Converter
{
    public class JsonToExcelConverter: ExcelHandlerBase<JsonToExcelConverterOptions, JsonToExcelConverterResult>
    {
        public override JsonToExcelConverterResult Process(JsonToExcelConverterOptions options)
        {
            Options = options;

            if (!options.Validate())
            {
                throw new ExcelToolsException("Wrong options");
            }

            try
            {
                Convert();

                return Result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        public void Convert()
        {
            using var reader = new JsonTextReader(new StreamReader(Options.FilePath));
            using var workbook = new XLWorkbook();

            var worksheet = workbook.Worksheets.Add("Data");
            var jsonToken = JToken.ReadFrom(reader);

            var levels = new Dictionary<string, int>();

            GetInformationAboutNesting(jsonToken, levels, 1);
                
            var currentColumn = 1;
                
            WriteJsonToWorksheet(jsonToken, worksheet, levels, ref currentColumn);
              
            ExcelHelper.DeleteEmptyColumns(worksheet);

            workbook.SaveAs(Options.ResultFilePath);
        }

        /// <summary>
        /// Рекурсивно чтение файла json
        /// </summary>
        /// <param name="token"></param>
        /// <param name="worksheet"></param>
        /// <param name="levels"></param>
        /// <param name="currentColumn"></param>
        /// <param name="currentLevel"></param>
        protected void WriteJsonToWorksheet(JToken token, IXLWorksheet worksheet, Dictionary<string, int> levels, ref int currentColumn, int currentLevel = 1)
        {
            switch (token)
            {
                case JObject jObject:
                    var properties = jObject.Properties().ToList();

                    foreach (var property in properties)
                    {
                        levels.TryGetValue(property.Name, out var currentRow);
                        worksheet.Cell(currentRow, currentColumn).Value = property.Name;
                        WriteJsonToWorksheet(property.Value, worksheet, levels, ref currentColumn, currentLevel + 1);

                        currentColumn++;
                    }
                    break;

                case JArray jArray:
                    if (jArray.All(item => item is JObject))
                    {
                        WriteArrayToWorksheet(jArray, worksheet, levels, currentLevel, currentColumn);
                    }
                    else
                    {
                        foreach (var t in jArray)
                        {
                            WriteJsonToWorksheet(t, worksheet, levels, ref currentColumn, currentLevel);
                        }
                    }

                    break;

                default:
                    var row = worksheet.Column(currentColumn).LastCellUsed()?.Address.RowNumber + 1 ?? 1;
                    worksheet.Cell(row, currentColumn).Value = token.ToString();

                    break;
            }
        }

/// <summary>
/// Чтение массива объктов json 
/// </summary>
/// <param name="jArray"></param>
/// <param name="worksheet"></param>
        protected void WriteArrayToWorksheet(JArray jArray, IXLWorksheet worksheet, Dictionary<string, int> levels, int currentLevel, int currentColumn)
        {
            if (jArray.Count == 0) return;
            
            foreach (var t in jArray)
            {
                var newLevels = levels.ToDictionary(entry => entry.Key, entry => entry.Value);

                var i = worksheet.LastRowUsed() != null ? worksheet.LastRowUsed().RowNumber() : 0;

                newLevels = ChangedInformationAboutNesting(newLevels, i);

                WriteJsonToWorksheet(t, worksheet, newLevels, ref currentColumn, currentLevel);

                currentColumn = 1;
            }
        }

        /// <summary>
        /// Получение информации о вложенности 
        /// </summary>
        /// <param name="token"></param>
        /// <param name="dictionary"></param>
        /// <param name="level"></param>
        protected void GetInformationAboutNesting(JToken token, Dictionary<string, int> dictionary, int level)
        {
            switch (token)
            {
                case JObject jObject:
                    foreach (var property in jObject.Properties())
                    {
                        dictionary[property.Name] = level;

                        GetInformationAboutNesting(property.Value, dictionary, level + 1);
                    }
                    break;

                case JArray jArray:
                    foreach (var t in jArray)
                    {
                        GetInformationAboutNesting(t, dictionary, level);
                    }
                    break;
            }
        }

        /// <summary>
        /// Изменение информации о вложенности
        /// </summary>
        /// <param name="levels"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        protected Dictionary<string, int> ChangedInformationAboutNesting(Dictionary<string, int> levels, int i)
        {
                foreach (var key in levels.Keys.ToList())
                {
                    levels[key] += i; 
                }

                return levels; 
        }
    }
}
