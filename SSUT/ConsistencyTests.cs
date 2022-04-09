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
    public  class ConsistencyTests : BaseTest
    {
        public const int LOOPS = 10000;

        private void LoopUsingSameDataSet(BaseTest context,Action r)
        {
            context.importer = this.importer;
            Parallel.For(0, LOOPS, new ParallelOptions { }, (i) =>
            {
                r();
            });
        }


        private void LoopUsingNewDataSet<TProp>(Action<TProp> r) where TProp : BaseTest, new()
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
                        TProp context = new TProp();
                        context.importer = importer;
                        r(context);

                    }
                }
            });

        }



        [Test]
        public void FullRangeCellTestLoadTheSameDataAgain()
        {
            SpreadSheetInstructionBuilderTest b = new SpreadSheetInstructionBuilderTest();
            LoopUsingSameDataSet(b, b.FullRangeCellTest);
        }
            

        [Test]
        public void FullRangeCellTestBurnInTest()
        {
            LoopUsingNewDataSet<SpreadSheetInstructionBuilderTest>(x => x.FullRangeCellTest());
        }
        [Test]
        public void MultipleSheetsBundlerTestBurnInTest()
        {

        }
    }
}
