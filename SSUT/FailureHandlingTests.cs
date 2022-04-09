using NUnit.Framework;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
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

        //[Test]
        //public void CloseConnectionToEarly()
        //{
        //    ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
        //}
    }
}
