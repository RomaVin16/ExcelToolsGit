using ClosedXML.Excel;
using ExcelTools.Abstraction;
using ExcelTools.Exceptions;

namespace ExcelTools.Cleaner
{
    public class Cleaner: ExcelHandlerBase<CleanOptions, CleanResult>
    {
        public override CleanResult Process(CleanOptions options)
        {
            Options = options;
           
            if (!options.Validate())
            {
                throw new ExcelToolsException("Wrong options");
            }

            try
            {
                DeleteEmptyStringInWorkbook();
                return Result;
            }
            catch (Exception e)
            {
                return ErrorResult(e.Message);
            }
        }

        /// <summary>
        /// удаление пустых строк в Excel файле
        /// </summary>
        protected void DeleteEmptyStringInWorkbook()
        {
            using (var workbook = new XLWorkbook(Options.FilePath))
            {
                DeleteEmptyRowsInSheet(workbook.Worksheet(Options.SheetNumber));
                
                workbook.SaveAs(Options.ResultFilePath);
            }
        }

        /// <summary>
        /// Удаление только пустых строк в конкретном листе
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        protected void DeleteEmptyRowsInSheet(IXLWorksheet item)
        {
            if (item.IsEmpty()) 
                return;

            for (var i = item.LastRowUsed().RowNumber(); i >= 1; i--)
            {
                Result.RowsProcessed++;
                if (item.Row(i).IsEmpty())
                {
                    item.Row(i).Delete();
                    Result.RowsRemoved++;
                }
            }
        }
    }
    }
