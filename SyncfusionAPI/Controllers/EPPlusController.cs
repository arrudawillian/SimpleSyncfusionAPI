using SyncfusionAPI.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using System.IO;
using Syncfusion.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SyncfusionAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EPPlusController : ControllerBase
    {
        /// <summary>
        /// Importação
        /// </summary>
        /// <param name="formFile"></param>
        /// <returns></returns>
        [HttpPost("import")]
        public DemoResponse<List<UserInfo>> Import(IFormFile formFile)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                return DemoResponse<List<UserInfo>>.GetResult(-1, "formfile is empty");
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return DemoResponse<List<UserInfo>>.GetResult(-1, "Not Support file extension");
            }

            var list = new List<UserInfo>();

            using (var stream = new MemoryStream())
            {
                formFile.CopyTo(stream);

                using (ExcelEngine excelEngine = new ExcelEngine())
                {

                    IApplication application = excelEngine.Excel;

                    application.DefaultVersion = ExcelVersion.Excel2016;

                    stream.Position = 0;

                    //Loads or open an existing workbook
                    IWorkbook workbook = excelEngine.Excel.Workbooks.Open(stream);
                    IWorksheet worksheet = workbook.Worksheets[0];
                    
                    var rowCount = worksheet.Rows.Count();

                    for (int row = 1; row < rowCount; row++)
                    {
                        list.Add(new UserInfo
                        {
                            UserName = worksheet.Rows[row].Cells[0].Value.ToString().Trim(),
                            Age = int.Parse(worksheet.Rows[row].Cells[1].Value.ToString().Trim())
                        });
                    }
                }
            }

            // add list to db ..  
            // here just read and return  

            return DemoResponse<List<UserInfo>>.GetResult(0, "OK", list);
        }


    }
}
