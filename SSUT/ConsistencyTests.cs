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
        public const int LOOPS = 100000;
        public const int THREADS = 4;


        public static void TestConsistency(Action<int> a)
        {
            Task[] pool = new Task[THREADS];
            for(int j =0;j < pool.Length;j++)
            {
                int scoped = j;
                pool[j] = Task.Run(() =>
                {
                    for (int i = scoped; i < LOOPS; i+= THREADS)
                    {
                        a(i);
                    }
                });
            }

            for(int i = 0; i < pool.Length; i++)
            {
                pool[i].Wait();
            }
        }


        public void LoopUsingSameDataSet(BaseTest context,Action r)
        {
            context.importer = this.importer;
            TestConsistency((int i) =>
            {
                r();
            });
        }


        public void LoopUsingNewDataSet<TProp>(Action<int,TProp> r) where TProp : BaseTest, new()
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
            TestConsistency((int i) =>
            {
                using (MemoryStream ms = new MemoryStream(data, false))
                {
                    using (IExcelImporter importer = new ExcelImporter(ms))
                    {
                        TProp context = new TProp();
                        context.importer = importer;
                        r(i,context);
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

        [Category(SKIP_SETUP)]
        [Test]
        public void FullRangeCellTestBurnInTest()
        {
            LoopUsingNewDataSet<SpreadSheetInstructionBuilderTest>((i,x) => x.FullRangeCellTest());
        }
        [Category(SKIP_SETUP)]
        [Test]
        public void MultipleSheetsBundlerTestBurnInTest()
        {
            LoopUsingNewDataSet<ConcurrencyTests>((i,x) => x.MultipleSheetsBundlerTest());
        }
    }
}
