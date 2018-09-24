using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Icris.Excel2Api
{
    public class ExcelReader
    {
        ExcelPackage excelPackage;
        string filename;

        public ExcelReader(string filename)
        {
            this.filename = filename;
            this.excelPackage = new ExcelPackage(new FileInfo(filename), true);
        }

        public List<Dictionary<string,object>> DataRecords
        {
            get
            {
                List<Dictionary<string, object>> value = new List<Dictionary<string, object>>();
                List<string> fields = new List<string>();
                var sheet = excelPackage.Workbook.Worksheets[1];
                var field = sheet.Cells[1, 1].Value;
                
                var nextcell = 2;
                while (field != null)
                {
                    fields.Add(field.ToString());
                    field = sheet.Cells[1, nextcell].Value;
                    nextcell++;
                }
                var row = 2;
                var firstcellvalue = sheet.Cells[row, 1].Value;
                while (firstcellvalue != null)
                {
                    Dictionary<string, object> record = new Dictionary<string, object>();
                    foreach(var column in fields)
                    {
                        record[column] = sheet.Cells[row, fields.IndexOf(column) + 1].Value;
                    }
                    value.Add(record);
                    row++;
                    firstcellvalue = sheet.Cells[row, 1].Value;
                }


                return value;
            }
        }
    }
}
