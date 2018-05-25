using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Icris.Excel2Api
{
    /// <summary>
    /// Exposes an excelsheet's logic
    /// </summary>
    public class ExcelCalculator
    {
        Workbook workbook;
        /// <summary>
        /// The input/outpu model of the excel sheet
        /// </summary>
        public Model Model { get; private set; }

        public ExcelCalculator(Workbook excel)
        {
            this.workbook = excel;
            ExtractModel();
        }

        /// <summary>
        /// Set an input value by key/value.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool SetInput(string key, object value)
        {
            var input = this.Model.Inputs[key];
            var inputsheet = (Worksheet)this.workbook.Sheets["Input"];
            inputsheet.Cells[input.Row, 4] = value;
            this.Model.Inputs[key].Value = value;
            var valid = (bool)inputsheet.get_Range($"E{input.Row}").Value;
            this.Model.Inputs[key].Valid = valid;
            var enabled = (bool)inputsheet.get_Range($"F{input.Row}").Value;
            this.Model.Inputs[key].Enabled = enabled;
            return valid;
        }



        void ExtractModel()
        {
            Model model = new Model();
            var inputsheet = (Worksheet)this.workbook.Sheets["Input"];
            var inputrow = 2;
            var input = (string)(inputsheet.get_Range("A2")).Value;
            while (!string.IsNullOrWhiteSpace(input))
            {
                var optionsvalue = (string)(inputsheet.get_Range($"G{inputrow}")).Value;
                model.Inputs.Add(input, new Input()
                {
                    Name = input,
                    Description = (string)(inputsheet.get_Range($"B{inputrow}")).Value,
                    Unit = (string)(inputsheet.get_Range($"C{inputrow}")).Value,
                    Value = (inputsheet.get_Range($"D{inputrow}")).Value,
                    Valid = (bool)(inputsheet.get_Range($"E{inputrow}")).Value,
                    Enabled = (bool)(inputsheet.get_Range($"F{inputrow}")).Value,
                    Options = string.IsNullOrEmpty(optionsvalue) ? new List<string>() : new List<string>(optionsvalue.Split(',').Select(x => x.Trim())),
                    Errormessage = (string)(inputsheet.get_Range($"H{inputrow}")).Value,
                    Row = inputrow
                });
                inputrow++;
                input = (string)(inputsheet.get_Range($"A{inputrow}")).Value;
            }

            var outputsheet = (Worksheet)this.workbook.Sheets["Output"];
            var output = (string)(outputsheet.get_Range("A2")).Value;
            var outputrow = 2;
            while (!string.IsNullOrWhiteSpace(output))
            {
                model.Outputs.Add(output, new Output()
                {
                    Name = output,
                    Description = (string)(outputsheet.get_Range($"B{outputrow}")).Value,
                    Value = (outputsheet.get_Range($"C{outputrow}")).Value,
                    Unit = (string)(outputsheet.get_Range($"D{outputrow}")).Value,
                    Row = outputrow
                });
                outputrow++;
                output = (string)(outputsheet.get_Range($"A{outputrow}")).Value;
            }

            this.Model = model;
        }
        /// <summary>
        /// Get an output value by key
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public object GetOutput(string key)
        {
            var output = this.Model.Outputs[key];
            var outputsheet = (Worksheet)this.workbook.Sheets["Output"];
            var value = outputsheet.get_Range($"C{output.Row}").Value;
            return value;
        }
        /// <summary>
        /// Create a Swagger JObject (serializable as a swagger.json file for automatic swagger-ui consumption)
        /// </summary>
        /// <param name="prefix"></param>
        /// <returns></returns>
        public JObject ToSwagger(string prefix)
        {
            Assembly assem = Assembly.GetAssembly(this.GetType());
            var resources = assem.GetManifestResourceNames();
            JObject swaggerdoc;
            JObject swaggerpath;
            using (var reader = new StreamReader(assem.GetManifestResourceStream("Icris.Excel2Api.swagger.swagger.json")))
            {
                swaggerdoc = JObject.Parse(reader.ReadToEnd());
            }
            using (var reader = new StreamReader(assem.GetManifestResourceStream("Icris.Excel2Api.swagger.path.json")))
            {
                swaggerpath = JObject.Parse(reader.ReadToEnd());
            }
            var inputdef = InputsToSwaggerDefinition();
            var outputdef = OutputsToSwaggerDefinition();
            swaggerdoc["definitions"] = new JObject();
            swaggerdoc["definitions"]["input"] = inputdef;
            swaggerdoc["definitions"]["output"] = outputdef;
            swaggerdoc["paths"] = new JObject();
            JObject inputref = new JObject();
            inputref["$ref"] = "#/definitions/input";
            JObject inputresponse = new JObject();
            inputresponse["200"] = JObject.FromObject(new
            {
                description = "Validation result",
                schema = inputref
            });
            JObject outputref = new JObject();
            outputref["$ref"] = "#/definitions/output";
            JObject outputresponse = new JObject();
            outputresponse["200"] = JObject.FromObject(new
            {
                description = "Calculation result",
                schema = outputref
            });
            swaggerdoc["paths"][$"/{prefix}/input"] = JObject.FromObject(new
            {
                post = new
                {
                    summary = "post input values to validate them against the excel sheet",
                    consumes = new string[] { "application/json" },
                    produces = new string[] { "application/json" },
                    parameters = new object[]
                    {
                        new
                        {
                            @in="body",
                            name="body",
                            description ="values of input to validate",
                            required="true",
                            schema = inputref
                        }
                    },
                    responses = inputresponse
                },
                get = new
                {
                    summary = "get the input values that are set in the excel sheet",
                    produces = new string[] { "application/json" },
                    responses = inputresponse
                }
            });
            swaggerdoc["paths"][$"/{prefix}/output"] = JObject.FromObject(new
            {
                post = new
                {
                    summary = "post input values to calculate output",
                    consumes = new string[] { "application/json" },
                    produces = new string[] { "application/json" },
                    parameters = new object[]
                    {
                        new
                        {
                            @in="body",
                            name="body",
                            description ="values of input to validate",
                            required="true",
                            schema = inputref
                        }
                    },
                    responses = outputresponse
                },
                get = new
                {
                    summary = "get the output object structure from the excel sheet",
                    produces = new string[] { "application/json" },
                    responses = outputresponse
                }
            });

            return swaggerdoc;
        }

        JObject OutputsToSwaggerDefinition()
        {

            JObject def = new JObject();
            def["type"] = "object";
            def["properties"] = new JObject();
            foreach (var kv in Model.Outputs)
            {
                JObject prop = new JObject();
                var type = kv.Value.Value.GetType();
                string valtype;
                switch (type.ToString())
                {
                    case "System.Int32":
                    case "System.Int64":
                        valtype = "integer";
                        break;
                    case "System.Boolean":
                        valtype = "boolean";
                        break;
                    case "System.Double":
                        valtype = "number";
                        break;
                    default:
                        valtype = "string";
                        break;
                }
                //numeric or integer or boolean
                def["properties"][kv.Key] = JObject.FromObject(new
                {
                    type = "object",
                    properties = new
                    {
                        value = new
                        {
                            type = valtype
                        },
                        unit = new
                        {
                            type = "string"
                        },
                        description = new
                        {
                            type = "string"
                        }
                    }
                });
            }
            return def;
        }
        JObject InputsToSwaggerDefinition()
        {
            JObject def = new JObject();
            def["type"] = "object";
            def["properties"] = new JObject();
            foreach (var kv in Model.Inputs)
            {
                JObject prop = new JObject();
                var type = kv.Value.Value.GetType();
                string valtype;
                switch (type.ToString())
                {
                    case "System.Int32":
                    case "System.Int64":
                        valtype = "integer";
                        break;
                    case "System.Boolean":
                        valtype = "boolean";
                        break;
                    case "System.Double":
                        valtype = "number";
                        break;
                    default:
                        valtype = "string";
                        break;
                }
                //numeric or integer or boolean
                def["properties"][kv.Key] = JObject.FromObject(new
                {
                    type = "object",
                    properties = new
                    {
                        value = new
                        {
                            type = valtype
                        },
                        valid = new
                        {
                            type = "boolean"
                        },
                        enabled = new
                        {
                            type = "boolean"
                        },
                        options = new
                        {
                            type = "array",
                            items = new
                            {
                                type = "string"
                            }
                        },
                        unit = new
                        {
                            type = "string"
                        },
                        errormessage = new
                        {
                            type = "string"
                        }
                    }
                });


            }
            return def;
        }
    }
}
