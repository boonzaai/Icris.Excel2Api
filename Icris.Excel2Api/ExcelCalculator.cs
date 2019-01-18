using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace Icris.Excel2Api
{
    /// <summary>
    /// Exposes an excelsheet's logic
    /// </summary>
    public class ExcelCalculator
    {
        ExcelPackage excelPackage;
        bool saveedits;
        string filename;

        //Workbook workbook;
        /// <summary>
        /// The input/outpu model of the excel sheet
        /// </summary>
        public Model Model { get; private set; }

        /// <summary>
        /// Create an excelcalculator based on a the given template file.
        /// </summary>
        /// <param name="filename"></param>
        public ExcelCalculator(string filename, bool saveedits = false)
        {
            try
            {
                //!!TODO: Locking mechanism !!
                this.filename = filename;
                while (!File.Exists(filename))
                    Thread.Sleep(100);
                this.excelPackage = new ExcelPackage(new FileInfo(filename), true);
                this.saveedits = saveedits;
                //this.workbook = excel;
                ExtractModel();
            }
            catch (Exception e)
            {
                Console.WriteLine("Failing creating excelcalculator: " + e.Message);
            }
        }
        public void Save()
        {
            try
            {
                this.excelPackage.SaveAs(new FileInfo(this.filename));
            }
            catch (Exception e)
            {
                Console.WriteLine("Error saving file. Will just continue now. Sorry. " + e.Message);
            }
        }

        public void Validate()
        {
            Thread.Sleep(100);
            try
            {
                var inputsheet = this.excelPackage.Workbook.Worksheets["Input"];
                this.excelPackage.Workbook.Calculate();
                foreach (var kv in this.Model.Inputs)
                {
                    this.Model.Inputs[kv.Key].Value = inputsheet.Cells[kv.Value.Row, 4].Value; // inputsheet.GetValue(kv.Value.Row, 3); // (inputsheet.get_Range($"D{kv.Value.Row}")).Value;
                    if (inputsheet.Cells[kv.Value.Row, 5].Value.GetType().Name == "ExcelErrorValue")
                        this.Model.Inputs[kv.Key].Valid = false;
                    else
                        this.Model.Inputs[kv.Key].Valid = (bool)(Convert.ChangeType(inputsheet.Cells[kv.Value.Row, 5].Value == null ? 0 : inputsheet.Cells[kv.Value.Row, 5].Value, typeof(bool)));// get_Range($"E{kv.Value.Row}")).Value;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to validate (dat rijmt ook nog) " + e.Message);
            }
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
            var inputsheet = this.excelPackage.Workbook.Worksheets["Input"];

            var valid = false;
            if (inputsheet.Cells[input.Row, 5].Value.GetType().Name != "ExcelErrorValue")
                valid = (bool)Convert.ChangeType(inputsheet.Cells[input.Row, 5].Value == null ? 0 : inputsheet.Cells[input.Row, 5].Value, typeof(bool)); // $"E{input.Row}").Value;
            var enabled = (bool)Convert.ChangeType(inputsheet.Cells[input.Row, 6].Value == null ? 0 : inputsheet.Cells[input.Row, 6].Value, typeof(bool)); //.get_Range($"F{input.Row}").Value;

            //refresh the model
            //TODO: Seeif this works....
            //if (enabled)
            inputsheet.SetValue(input.Row, 4, value); //.Cells[input.Row, 4] = value;

            this.excelPackage.Workbook.Calculate();


            //this.Model.Inputs[key].Value = value;
            this.Model.Inputs[key].Valid = inputsheet.GetValue<bool>(input.Row, 5);
            //var enabled = (bool)inputsheet.get_Range($"F{input.Row}").Value;
            //this.Model.Inputs[key].Enabled = enabled;
            //if (saveedits)
            //{
            //    this.excelPackage.SaveAs(new FileInfo(this.filename));
            //}
            return this.Model.Inputs[key].Valid;
        }




        void ExtractModel()
        {
            Model model = new Model();
            var inputsheet = this.excelPackage.Workbook.Worksheets["Input"];
            //var inputsheet = (Worksheet)this.workbook.Sheets["Input"];
            var inputrow = 2;
            var input = (string)inputsheet.Cells[inputrow, 1].Value;
            //var input = (string)(inp); // .get_Range("A2")).Value;
            while (!string.IsNullOrWhiteSpace(input))
            {
                var optionsvalue = (string)(inputsheet.Cells[inputrow, 8].Value);// get_Range($"G{inputrow}")).Value;
                bool valid = false;
                if (inputsheet.Cells[inputrow, 5].Value.GetType().Name != "ExcelErrorValue")
                    valid = (bool)Convert.ChangeType(inputsheet.Cells[inputrow, 5].Value == null ? 0 : inputsheet.Cells[inputrow, 5].Value, typeof(bool));
                bool enabled = (bool)Convert.ChangeType(inputsheet.Cells[inputrow, 6].Value == null ? 0 : inputsheet.Cells[inputrow, 6].Value, typeof(bool));
                bool visible = (bool)Convert.ChangeType(inputsheet.Cells[inputrow, 7].Value == null ? 0 : inputsheet.Cells[inputrow, 7].Value, typeof(bool));

                model.Inputs.Add(input, new Input()
                {
                    Name = input,
                    Description = (string)(inputsheet.Cells[inputrow, 2].Value),   // get_Range($"B{inputrow}")).Value,
                    Unit = (string)(inputsheet.Cells[inputrow, 3].Value),           //get_Range($"C{inputrow}")).Value,
                    Value = (inputsheet.Cells[inputrow, 4].Value),                  //get_Range($"D{inputrow}")).Value,
                    Valid = valid,           //get_Range($"E{inputrow}")).Value,
                    Enabled = enabled,         //get_Range($"F{inputrow}")).Value,
                    Visible = visible,
                    Options = string.IsNullOrEmpty(optionsvalue) ? new List<string>() : new List<string>(optionsvalue.Split(',').Select(x => x.Trim())),
                    Errormessage = (string)(inputsheet.Cells[inputrow, 9].Value),  //get_Range($"H{inputrow}")).Value,
                    Row = inputrow
                });
                inputrow++;
                input = (string)inputsheet.Cells[inputrow, 1].Value;             // get_Range($"A{inputrow}")).Value;
            }

            //var outputsheet = (Worksheet)this.workbook.Sheets["Output"];
            var outputsheet = this.excelPackage.Workbook.Worksheets["Output"];
            var outputrow = 2;
            var output = (string)(outputsheet.Cells[outputrow, 1].Value);     //.get_Range("A2")).Value;
            while (!string.IsNullOrWhiteSpace(output))
            {
                model.Outputs.Add(output, new Output()
                {
                    Name = output,
                    Description = (string)(outputsheet.Cells[outputrow, 2].Value),  //get_Range($"B{outputrow}")).Value,
                    Value = (outputsheet.Cells[outputrow, 3].Value),                //get_Range($"C{outputrow}")).Value,
                    Unit = (string)(outputsheet.Cells[outputrow, 4].Value),         //get_Range($"D{outputrow}")).Value,
                    Row = outputrow
                });
                outputrow++;
                output = (string)(outputsheet.Cells[outputrow, 1].Value);          //.get_Range($"A{outputrow}")).Value;
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
            var outputsheet = this.excelPackage.Workbook.Worksheets["Output"];
            //outputsheet.Cells[output.Row, 3].Formula = outputsheet.Cells.Formula.Replace(';', ',');
            //outputsheet.Calculate();
            outputsheet.Cells[output.Row, 3].Calculate(new OfficeOpenXml.FormulaParsing.ExcelCalculationOption() { AllowCirculareReferences = true });
            var value = outputsheet.Cells[output.Row, 3].Value;
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
