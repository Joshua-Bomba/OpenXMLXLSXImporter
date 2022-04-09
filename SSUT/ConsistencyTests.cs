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
        public const int THRADS = 8;


        public static void TestConsistency(Action a)
        {
            //this method is alot easier to track when things freeze the the ParallerFor
            Task[] pool = new Task[THRADS];
            int poolOffset = 0;
            for (int i = 0; i < LOOPS; i++)
            {
                Task t = Task.Run(a);
                if (i < pool.Length)
                {
                    pool[i] = t;
                }
                else
                {
                    if (poolOffset >= pool.Length)
                    {
                        poolOffset = 0;
                    }
                    pool[poolOffset].Wait();
                    pool[poolOffset] = t;
                    poolOffset++;
                }
            }
            for(int i =0;i < pool.Length;i++)
            {
                if(pool[i] != null)
                {
                    pool[i].Wait();
                }
            }
        }


        private void LoopUsingSameDataSet(BaseTest context,Action r)
        {
            context.importer = this.importer;
            TestConsistency(() =>
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
            TestConsistency(() =>
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
            LoopUsingNewDataSet<ConcurrencyTests>(x => x.MultipleSheetsBundlerTest());
        }
    }
}
