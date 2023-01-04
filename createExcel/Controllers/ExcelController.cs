using createExcel.API;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace createExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {


        public static excel excel = new excel();
        [HttpPut]
        [Route("addExcel")]

        public async Task<IActionResult> addExcelAsync(string excelFile, string sheetName, string cell, string value)
        {

            string code = await excel.addExcel( excelFile,  sheetName,  cell,  value);
            if (string.IsNullOrEmpty(code))
            {
                return BadRequest();
            }
            else
            {
                return Ok(code);
            }
        }
    }
}
