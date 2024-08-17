using ClosedXML.Excel;
using ExcelTools.Abstraction;

namespace ExcelTools.Cleaner
{
    public class Cleaner: ExcelHandlerBase<CleanOptions, CleanResult>
    {
        private CleanOptions options = new CleanOptions();
        private CleanResult result  = new CleanResult();


        public override CleanResult Process(CleanOptions options)
        {
            this.options = options;
           
            if (!options.Validate())
            {
                return ErrorResult("Wrong options");
            }

            try
            {
                DeleteEmptyStringInWorkbook();
                return this.result;
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
            using (var workbook = new XLWorkbook(options.FilePath))
            {
                foreach (var item in workbook.Worksheets)
                {
                    DeleteEmptyRowsInSheet(item);
                }

                workbook.SaveAs(options.ResultFilePath);
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
                result.RowsProcessed++;
                if (item.Row(i).IsEmpty())
                {
                    item.Row(i).Delete();
                    result.RowsRemoved++;
                }
            }
        }

   

       
    }
    }
