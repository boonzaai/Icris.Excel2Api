//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace Icris.Excel2Api.Console.controllers
{
    
    public class SwaggerController:ApiController
    {
        [HttpGet]        
        public IHttpActionResult Get()
        {
            var path = this.ActionContext.RequestContext.RouteData.Values["path"].ToString();
                        
            var definition = new ExcelCalculator(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
                + $"\\sheets\\{path}.xlsx").ToSwagger(path);

            return Ok(definition);            
        }
    }
}
