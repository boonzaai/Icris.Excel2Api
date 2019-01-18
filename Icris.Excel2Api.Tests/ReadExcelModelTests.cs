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
        string tempfile = Path.GetTempFileName();

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
        }
        void Cleanup()
        {
        }
        [TestMethod]
        public void TestTestModelExtraction()
        {
            GetTestWorkbook();
            var calculator = new ExcelCalculator(tempfile);
            var model = calculator.Model;
            Assert.AreEqual(4, model.Inputs.Count);
            Assert.AreEqual(4, model.Outputs.Count);
            Cleanup();
        }

        [TestMethod]
        public void TestTestModelValidation()
        {
            GetTestWorkbook();
            var model = new ExcelCalculator(tempfile);

            Assert.AreEqual(false, model.SetInput("Width", 12));
            
            Assert.AreEqual(true, model.SetInput("Length", 8));
            Cleanup();
        }


        [TestMethod]
        public void TestTestModelCalculation()
        {
            GetTestWorkbook();
            var model = new ExcelCalculator(tempfile);
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
            var model = new ExcelCalculator(tempfile);

            Assert.IsNotNull(model.ToSwagger("test"));
            Cleanup();
        }
        [TestMethod]
        public void TestTestAdvancedFormula()
        {
            GetTestWorkbook();
            var model = new ExcelCalculator(tempfile);
            model.SetInput("Width", 2);
            model.SetInput("Height", 2);
            model.SetInput("Length", 2);
            model.SetInput("Material", "Aluminum");

            Assert.AreEqual(0.0216, model.GetOutput("Weight"));
        }
    }
}
