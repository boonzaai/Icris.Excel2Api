using Newtonsoft.Json.Linq;
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
    public class SheetController : ApiController
    {
        [HttpGet]
        public IHttpActionResult Get()
        {
            var path = this.ActionContext.RequestContext.RouteData.Values["path"].ToString();
            if (path.ToLower() == "sheets")
            {
                var files = Directory.EnumerateFiles(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\sheets", "*.xlsx");
                return Ok(JArray.FromObject(files.Select(x => Path.GetFileNameWithoutExtension(x))));
            }

            var sheet = path.Split('/')[0];
            var action = path.Split('/')[1];

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            var wbs = excel.Workbooks;
            var wb = wbs.Open(
                Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
                + $"\\sheets\\{sheet}.xlsx");
            var definition = new JObject();
            //new Model()
            //new ExcelCalculato

            //var path = this.ActionContext.RequestContext.RouteData.Values["path"].ToString();
            switch (action.ToLower())
            {
                case "input":
                    new ExcelCalculator(wb).Model.Inputs.Select(y => y.Value).ToList().ForEach(x =>
                    {
                        definition[x.Name] = JObject.FromObject(new
                        {
                            value = x.Value,
                            valid = x.Valid,
                            enabled = x.Enabled,
                            options = x.Options,
                            unit = x.Unit,
                            errormessage = x.Errormessage
                        });
                    });
                    break;
                case "output":
                    new ExcelCalculator(wb).Model.Outputs.Select(y => y.Value).ToList().ForEach(x =>
                    {
                        definition[x.Name] = JObject.FromObject(new
                        {
                            value = x.Value,
                            description = x.Description,
                            unit = x.Unit
                        });
                    });
                    break;
            }
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
            GC.Collect();
            return Ok(definition);
        }
        [HttpPost]
        public IHttpActionResult Post(JObject payload)
        {
            var path = this.ActionContext.RequestContext.RouteData.Values["path"].ToString();
            if (path.ToLower() == "sheets")
            {
                var files = Directory.EnumerateFiles(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\sheets", "*.xlsx");
                return Ok(JArray.FromObject(files.Select(x => Path.GetFileNameWithoutExtension(x))));
            }

            var sheet = path.Split('/')[0];
            var action = path.Split('/')[1];

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            var wbs = excel.Workbooks;
            var wb = wbs.Open(
                Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
                + $"\\sheets\\{sheet}.xlsx");
            var definition = new JObject();
            //new Model()
            //new ExcelCalculato

            var calculator = new ExcelCalculator(wb);
            //1. Set the input
            payload.Properties().ToList().ForEach(x =>
            {
                var oldvalue = calculator.Model.Inputs[x.Name].Value;
                var val = x.Value["value"];
                var converted = Convert.ChangeType(val, oldvalue.GetType());
                calculator.SetInput(x.Name, converted);
            });

            switch (action.ToLower())
            {
                //Case input validation
                case "input":
                    //2. Fetch the validated inputs
                    calculator.Model.Inputs.Select(y => y.Value).ToList().ForEach(x =>
                    {
                        definition[x.Name] = JObject.FromObject(new
                        {
                            value = x.Value,
                            valid = x.Valid,
                            unit = x.Unit,
                            errormessage = x.Errormessage
                        });
                    });
                    break;
                //Case output calculation
                case "output":
                    //2. Fetch calculated outputs

                    calculator.Model.Outputs.Select(y => y.Value).ToList().ForEach(x =>
                    {
                        definition[x.Name] = JObject.FromObject(new
                        {
                            value = calculator.GetOutput(x.Name),
                            description = x.Description
                        });
                    });
                    break;
            }
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
            GC.Collect();
            return Ok(definition);
        }
    }
}
