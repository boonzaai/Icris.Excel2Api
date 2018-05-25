using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Icris.Excel2Api.Tests
{
    [TestClass]
    public class ReadExcelModelTests
    {
        Application excel = new Microsoft.Office.Interop.Excel.Application();
        string tempfile = Path.GetTempFileName();
        Workbooks workbooks;
        Workbook workbook;

        void GetTestWorkbook()
        {
            Assembly assem = Assembly.GetAssembly(this.GetType());
            var resources = assem.GetManifestResourceNames();
            using (Stream stream = assem.GetManifestResourceStream("Icris.Excel2Api.Tests.data.test.xlsx"))
            {
                using (BinaryWriter target = new BinaryWriter(File.OpenWrite(tempfile)))
                {
                    using (var reader = new BinaryReader(stream))
                    {
                        target.Write(reader.ReadBytes((int)reader.BaseStream.Length));
                    }
                }
            }
            workbooks = excel.Workbooks;
            workbook = workbooks.Open(tempfile);
        }
        void Cleanup()
        {
            //Excel process is a quite persistent bugger, it won't die without a lot of fuss...
            workbook.Close(false);
            workbooks.Close();
            excel.Quit();
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(excel);
            workbook = null;
            workbooks = null;
            excel = null;
            GC.Collect();
        }
        [TestMethod]
        public void TestTestModelExtraction()
        {
            GetTestWorkbook();
            //var sheet = (Worksheet)wb.Sheets["Input"];
            //var val = (sheet.get_Range("A2")).Value;
            var calculator = new ExcelCalculator(workbook);
            var model = calculator.Model;
            Assert.AreEqual(4, model.Inputs.Count);
            Assert.AreEqual(4, model.Outputs.Count);

            Cleanup();
        }

        [TestMethod]
        public void TestTestModelValidation()
        {
            GetTestWorkbook();

            var model = new ExcelCalculator(workbook);
            Assert.AreEqual(false, model.SetInput("Width", 12));
            Assert.AreEqual(true, model.SetInput("Length", 8));

            Cleanup();
        }


        [TestMethod]
        public void TestTestModelCalculation()
        {
            GetTestWorkbook();

            var model = new ExcelCalculator(workbook);
            model.SetInput("Width", 3);
            model.SetInput("Height", 3);
            model.SetInput("Length", 3);

            Assert.AreEqual(27.0, model.GetOutput("Volume"));
            Assert.AreEqual(54.0, model.GetOutput("Area"));
            Cleanup();
        }

        [TestMethod]
        public void TestTestModelSwagger()
        {
            GetTestWorkbook();
            var model = new ExcelCalculator(workbook);

            Assert.IsNotNull(model.ToSwagger("test"));
            Cleanup();
        }
    }
}
