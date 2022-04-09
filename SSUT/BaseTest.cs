using NUnit.Framework;
using OpenXMLXLSXImporter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSUT
{
    public abstract class BaseTest
    {
        protected Stream stream = null;
        protected IExcelImporter importer = null;

        [SetUp]
        public void Setup()
        {
            stream = new MemoryStream();
            using (FileStream fs = File.OpenRead(Global.TEST_FILE))
            {
                fs.CopyTo(stream);
            }
            importer = new ExcelImporter(stream);
        }
        [TearDown]
        public void TearDown()
        {
            importer?.Dispose();
            stream?.Close();
            stream?.Dispose();
        }
    }
}
