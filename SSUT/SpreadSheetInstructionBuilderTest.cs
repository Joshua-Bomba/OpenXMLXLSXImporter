using NUnit.Framework;
using OpenXMLXLSXImporter;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SSUT
{
    public class SpreadSheetInstructionBuilderTest : BaseTest
    {
        
        [Test]
        public void LoadTitleCellTestUsingRunner()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ICellData result = builder.Runner.LoadSingleCell(1, "A").GetAwaiter().GetResult();
            Assert.IsTrue(result.Content() == "A Header");
        }
        [Test]
        public void LoadTitleCellTestUsingBundler()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ISpreadSheetInstruction instruction = builder.Bundler.LoadSingleCell(1, "A");
            builder.Bundler.BundleRequeset(new List<ISpreadSheetInstruction> { instruction }).GetAwaiter().GetResult();
            ICellData result = instruction.GetResults().FirstAsync().GetAwaiter().GetResult();
            Assert.IsTrue(result.Content() == "A Header");
        }

        [Test]
        public void LoadTitleCellTestUsingBundlerResults()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ISpreadSheetInstruction instruction = builder.Bundler.LoadSingleCell(1, "A");
            IAsyncEnumerable<ICellData> results = builder.Bundler.GetBundledResults(new List<ISpreadSheetInstruction> { instruction});
            ICellData result = results.FirstAsync().GetAwaiter().GetResult();
            Assert.IsTrue(result.Content() == "A Header");
        }
       
        [Test]
        public void TwoCells()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ISpreadSheetInstruction[] bundle = new ISpreadSheetInstruction[] {
                builder.Bundler.LoadSingleCell(2, "B"),
                builder.Bundler.LoadSingleCell(2, "A")
            };

            ICellData c = builder.Bundler.GetBundledResults(bundle).FirstAsync().GetAwaiter().GetResult();
             Assert.IsTrue(c.Content() == "Data in another cell");
        }
        
        [Test]
        public void RangeCellTest()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ValueTask<ICellData[]> columnRange = builder.Runner.LoadColumnRange(4, "A", "P").ToArrayAsync();
            ValueTask<ICellData[]> rowRange = builder.Runner.LoadRowRange("G", 1, 7).ToArrayAsync();
            Global.CheckResultsAsync(columnRange, rowRange).GetAwaiter().GetResult();
        }



        [Test]
        public void FullRangeCellTest()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ValueTask<ICellData[]> columnRange = builder.Runner.LoadFullColumnRange(4).ToArrayAsync();
            ValueTask<ICellData[]> rowRange = builder.Runner.LoadFullRowRange("G").ToArrayAsync();
            Global.CheckResultsAsync(columnRange, rowRange).GetAwaiter().GetResult();
        }

    }
}