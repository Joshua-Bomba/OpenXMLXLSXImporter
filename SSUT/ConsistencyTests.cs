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
        [Test]
        public void FullRangeCellTestLoadTheSameDataAgain()
        {
            SpreadSheetInstructionBuilderTest b = new SpreadSheetInstructionBuilderTest();
            b.importer = this.importer;
            Parallel.For(0, LOOPS, new ParallelOptions { }, (i) =>
            {
                try
                {
                    b.FullRangeCellTest();
                }
                catch
                {

                }
            });
        }

        [Test]
        public void FullRangeCellTestBurnInTest()
        {
            SpreadSheetInstructionBuilderTest b = new SpreadSheetInstructionBuilderTest();
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
                            b.importer = importer;
                            b.FullRangeCellTest();
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
