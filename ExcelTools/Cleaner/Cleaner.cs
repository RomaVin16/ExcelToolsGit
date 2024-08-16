using ClosedXML.Excel;
using ExcelTools.Abstraction;

namespace ExcelTools.Cleaner
{
    public class Cleaner: ExcelHandlerBase
    {
        private readonly CleanOptions cleanOptions;
        private readonly CleanResult cleanResult;

public Cleaner(CleanOptions cleanOptions, CleanResult cleanResult)
        {
            this.cleanOptions = cleanOptions;
            this.cleanResult = cleanResult;
        }

        /// <summary>
        /// Удаление только пустых строк в конкретном листе
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        public void DeleteEmptyRowsInSheet(IXLWorksheet item, CleanResult result)
        {
            if (item.IsEmpty()) return;

            for (var i = item.LastRowUsed().RowNumber(); i >= 1; i--)
            {
                result.RowsProcessed++;
                if (item.Row(i).IsEmpty())
                {
                    item.Row(i).Delete();
                    result.RowsRemoved++;
                }
            }
        }

        /// <summary>
        /// удаление пустых строк в Excel файле
        /// </summary>
        public CleanResult DeleteEmptyStringInWorkbook()
        {
            using (var workbook = new XLWorkbook(cleanOptions.FilePath))
            {
                foreach (var item in workbook.Worksheets)
                {
                    DeleteEmptyRowsInSheet(item, cleanResult);
                }

                workbook.SaveAs(cleanOptions.ResultFilePath);

                return cleanResult;
            }
        }

        public override void Process()
        {
            try
            {
                if (cleanOptions.FilePath != null)
                {
                    DeleteEmptyStringInWorkbook();
                }
                else
                {
                    throw new Exception("Настройки не заданы.");
                }
            }
            catch (Exception e)
            {
                cleanResult.Code = (ExcelResultBase.ResultCode)1;
                cleanResult.ErrorMessage = e.Message;
            }
        }
    }
    }
