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
        public const int LOOPS = 1000000;
        public const int THREADS = 12;


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


        public void LoopUsingSameDataSet<TProp>(TProp context,Action<int,TProp> r) where TProp : BaseTest
        {
            context.importer = this.importer;
            TestConsistency((int i) =>
            {
                r(i,context);
            });
        }

        public static byte[] GetData()
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
            return data;
        }


        public void LoopUsingNewDataSet<TProp>(Action<int,TProp> r) where TProp : BaseTest, new()
        {
            byte[] data = GetData();
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
            SpreadSheetInstructionBuilderTest s = new SpreadSheetInstructionBuilderTest();
            LoopUsingSameDataSet(s, (i, x) => x.FullRangeCellTest());
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
