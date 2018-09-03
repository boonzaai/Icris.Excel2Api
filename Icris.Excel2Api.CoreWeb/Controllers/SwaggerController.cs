//using Microsoft.Office.Interop.Excel;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace Icris.Excel2Api.CoreWeb.controllers
{

    public class SwaggerController : ControllerBase
    {
        [Route("api/swagger/{sheet}")]
        [HttpGet]
        public ActionResult Get(string sheet)
        {
            var definition = new ExcelCalculator(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
                + $"/sheets/{sheet}.xlsx").ToSwagger(sheet);

            return Ok(definition);
        }
    }
}
