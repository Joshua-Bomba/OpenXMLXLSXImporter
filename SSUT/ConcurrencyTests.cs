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
        public void MultipleSheetsBundlerTest()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ISpreadSheetInstructionBuilder builder2 = importer.GetSheetBuilder(Global.SHEET2).GetAwaiter().GetResult();
            ISpreadSheetInstructionBundler bundler1 =  builder.GetBundler();
            bundler1.LoadColumnRange(5, "B", "J");
            bundler1.LoadColumnRange(6, "B", "J");
            bundler1.LoadColumnRange(7, "B", "J");
            ISpreadSheetInstructionBundler bundler2 = builder2.GetBundler();
            bundler2.LoadRowRange("A", 1, 4);
            bundler2.LoadRowRange("B", 1, 4);
            
            string[] r1 = bundler1.GetBundledResults().ToArrayAsync().GetAwaiter().GetResult().Select(x => x.Content()).ToArray();
            string[] r2 = bundler2.GetBundledResults().ToArrayAsync().GetAwaiter().GetResult().Select(x => x.Content()).ToArray();
        }

        [Test]
        public void MultipleSheetInterwined()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(Global.SHEET1).GetAwaiter().GetResult();
            ISpreadSheetInstructionBuilder builder2 = importer.GetSheetBuilder(Global.SHEET2).GetAwaiter().GetResult();

            ValueTask<ICellData[]>[] bundle1 = new ValueTask<ICellData[]>[3];
            ValueTask<ICellData[]>[] bundle2 = new ValueTask<ICellData[]>[2];

            bundle1[0] = builder.Runner.LoadColumnRange(5, "B", "J").ToArrayAsync();
            bundle2[0] = builder2.Runner.LoadRowRange("A", 1, 4).ToArrayAsync();
            bundle1[1] = builder.Runner.LoadColumnRange(6, "B", "J").ToArrayAsync();
            bundle2[1] = builder2.Runner.LoadRowRange("B", 1, 4).ToArrayAsync();
            bundle1[2] = builder.Runner.LoadColumnRange(7, "B", "J").ToArrayAsync();

            List<string> r1 = new List<string>();
            List<string> r2 = new List<string>();
            ICellData[] r;
            var f = async () =>
            {
                for (int i = 0; i < bundle1.Length; i++)
                {
                    r = await bundle1[i];
                    foreach(ICellData c in r)
                    {
                        r1.Add(c.Content());
                    }
                    if (i < bundle2.Length)
                    {
                        r = await bundle2[i];
                        foreach (ICellData c in r)
                        {
                            r2.Add(c.Content());
                        }
                    }
                }
            };
            f().GetAwaiter().GetResult();

           

        }



      
    }
}
