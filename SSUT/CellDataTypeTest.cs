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
    public class CellDataTypeTest : BaseTest
    {
        [Test]
        public void EmptyCellTest()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ISpreadSheetInstructionBundler bundler = builder.GetBundler();
            bundler.LoadColumnRange(3, "A", "Z");
            bundler.LoadRowRange("A", 5, 15);
            bundler.LoadFullRowRange("K");
            bundler.LoadFullRowRange("L");
            bundler.LoadFullRowRange("M");
            bundler.LoadFullRowRange("N");

            ICellData[] result = bundler.GetBundledResults().ToArrayAsync().GetAwaiter().GetResult();
            Assert.IsFalse(result.Any(x => x is not EmptyCell));
        }
        [Test]
        public void OldReferenceAsEmptyCell()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ICellData oldReferenceCell = builder.Runner.LoadSingleCell(3, "D").GetAwaiter().GetResult();
            Assert.IsTrue(oldReferenceCell is EmptyCell);
        }
    }
}
