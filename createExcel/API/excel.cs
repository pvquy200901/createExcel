using IronXL;

namespace createExcel.API
{
    public class excel
    {
        public async Task<string> addExcel(string excelFile, string sheetName, string cell, string value)
        {
            try
            {

                WorkBook workBook = WorkBook.Load(@"" + excelFile);
                WorkSheet sheet = workBook.GetWorkSheet(sheetName);
                sheet[cell].Value = value;
                workBook.Save();
                return "Ok";
            }
            catch
            {
                return "";
            }
        }
    }
}
