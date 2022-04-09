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
    public class ConcurrencyTests : BaseTest
    {
        public const int LOOPS = 10000;

        [Test]
        public void MultipleImporters()
        {
            Task t = new Task(() =>
            {
                ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
                ICellData result = builder.Runner.LoadSingleCell(1, "A").GetAwaiter().GetResult();
                Assert.IsTrue(result.Content() == "A Header");
            });

            Task t2 = new Task(() =>
            {

                ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
                ICellData result = builder.Runner.LoadSingleCell(1, "A").GetAwaiter().GetResult();
                Assert.IsTrue(result.Content() == "A Header");
            });

            t.Start();
            t2.Start();

            t.Wait();
            t2.Wait();
        }



        [Test]
        public void FullRangeCellTestLoadTheSameDataAgain()
        {
            Parallel.For(0, LOOPS, new ParallelOptions { }, (i) =>
            {
                try
                {
                    ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
                    ValueTask<ICellData[]> columnRange = builder.Runner.LoadFullColumnRange(4).ToArrayAsync();
                    ValueTask<ICellData[]> rowRange = builder.Runner.LoadFullRowRange("G").ToArrayAsync();
                    Global.CheckResultsAsync(columnRange, rowRange).GetAwaiter().GetResult();
                }
                catch
                {

                }
            });
        }

        [Test]
        public void FullRangeCellTestBurnInTest()
        {
            byte[] data;
            using (MemoryStream baseStream = new MemoryStream())
            {
                using (FileStream fs = File.OpenRead(Global.TEST_FILE))
                {
                    fs.CopyTo(baseStream);
                }
                data = baseStream.ToArray();
            }

            Parallel.For(0, LOOPS, new ParallelOptions { MaxDegreeOfParallelism = 2 }, (i) =>
            {
                using (MemoryStream ms = new MemoryStream(data, false))
                {
                    using (IExcelImporter importer = new ExcelImporter(ms))
                    {
                        try
                        {
                            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
                            ValueTask<ICellData[]> columnRange = builder.Runner.LoadFullColumnRange(4).ToArrayAsync();
                            ValueTask<ICellData[]> rowRange = builder.Runner.LoadFullRowRange("G").ToArrayAsync();
                            Global.CheckResultsAsync(columnRange, rowRange).GetAwaiter().GetResult();
                        }
                        catch
                        {

                        }

                    }
                }
            });

        }
    }
}
