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
        public Stream stream = null;
        public IExcelImporter importer = null;
        public const string SKIP_SETUP = "SkipSetup";

        [SetUp]
        public void Setup()
        {
            bool skip = TestContext.CurrentContext.Test.Properties["Category"].Contains(SKIP_SETUP);
            if (!skip)
            {
                stream = new MemoryStream();
                using (FileStream fs = File.OpenRead(Global.TEST_FILE))
                {
                    fs.CopyTo(stream);
                }
                stream.Position = 0;
                importer = new ExcelImporter(stream);
            }
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
