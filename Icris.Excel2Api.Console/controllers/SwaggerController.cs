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


            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            var wbs = excel.Workbooks;
            var wb = wbs.Open(
                Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
                + $"\\sheets\\{path}.xlsx");
            var definition = new ExcelCalculator(wb).ToSwagger(path);
            //Excel process is a quite persistent bugger, it won't die without a lot of fuss...
            wb.Close(false);
            wbs.Close();
            excel.Quit();
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(wbs);
            Marshal.ReleaseComObject(excel);
            wb = null;
            wbs = null;
            excel = null;
            GC.Collect(); return Ok(definition);
            
            //return Ok(new Icris.Excel2Api.ExcelCalculator());
            //Excel
            //return Ok(Singleton.Instance.DataContainer.ToSwaggerDefinition());
        }
    }
}
