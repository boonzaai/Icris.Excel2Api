using System;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Icris.Excel2Api.Tests
{
    [TestClass]
    public class ExcelReaderTestscs
    {
        string tempfile = Path.GetTempFileName();

        void GetTestWorkbook()
        {
            Assembly assem = Assembly.GetAssembly(this.GetType());
            var resources = assem.GetManifestResourceNames();
            using (Stream stream = assem.GetManifestResourceStream("Icris.Excel2Api.Tests.data.datarecords.xlsx"))
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
        [TestMethod]
        public void ReadDataRecordsTest()
        {
            GetTestWorkbook();
            var result = new ExcelReader(tempfile).DataRecords;
        }
    }
}
