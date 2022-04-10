using NUnit.Framework;
using OpenXMLXLSXImporter;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSUT
{
    [TestFixture]
    public class FailureHandlingTests : BaseTest
    {
        [Test]
        public void NonExistingSheet()
        {
            try
            {
                ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder("not a real sheet").GetAwaiter().GetResult();
                ICellData result = builder.Runner.LoadSingleCell(1, "A").GetAwaiter().GetResult();
                Assert.Fail();
            }
            catch (Exception ex)
            {

            }

        }
        [Test]
        [Category(SKIP_SETUP)]
        public void SlowDataTest()
        {
            using (DelayedData d = new DelayedData())
            {
                using (FileStream fs = File.OpenRead(Global.TEST_FILE))
                {
                    fs.CopyTo(d);
                }
                using (IExcelImporter importer = new ExcelImporter(d))
                {
                    
                }
            }
            
        }

        //[Test]
        //public void CloseConnectionToEarly()
        //{
        //    ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
        //}
    }
}
