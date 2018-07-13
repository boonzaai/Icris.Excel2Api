# Icris.Excel2Api
A convention-based library that translates logic from excel sheets to Swagger documented API's.

Basically, this library reads inputs and outputs defined in an excel sheet (from an input resp. output tab). Each input is defined by a name, value, valid and errormessage value. These are exposed by a GET method on the excelsheet/input API. The same goes for the outputs defined in the output tab. Logic that calculates the 'valid' property of an input, and the outcome of an output, can be added to the excel via formula's. 
Based on the Excel definition, a Swagger UI is generated automatically including input/output definition.

![overview](https://raw.githubusercontent.com/boonzaai/Icris.Excel2Api/master/overview.png)

This library uses Excel interop to read a convention-based excel model and exposes this as a (Swagger documented) web api. Prerequisites are of course Excel. For more information, please review the unit tests & example console host project.
