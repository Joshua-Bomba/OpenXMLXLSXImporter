using NUnit.Framework;
using OpenXMLXLSXImporter.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSUT
{
    public class EncodingTest
    {
        private void TestRange(string[] testColumns)
        {
            uint[] indexs = new uint[testColumns.Length];
            for (uint i = 0; i < testColumns.Length; i++)
            {
                indexs[i] = ExcelColumnHelper.GetColumnStringAsIndex(testColumns[i]);
            }

            //Reconvert it back
            string[] unencoded = new string[indexs.Length];

            for (uint i = 0; i < indexs.Length; i++)
            {
                unencoded[i] = ExcelColumnHelper.GetColumnIndexAsString(indexs[i]);
            }

            //Ensure they are equal to the original value
            for (uint i = 0; i < unencoded.Length; i++)
            {
                Assert.IsTrue(testColumns[i] == unencoded[i]);
            }
        }

        [Test]
        public void IndexEncodingTest2Digits()
        {
            string[] testColumns = new string[] { "X", "Y", "Z", "AA", "AB", "AC" };
            TestRange(testColumns);
        }
        [Test]
        public void MoveSecondSymbolTest()
        {
            string[] testColumns = new string[] {  "AZ", "BA", "BB", "BC" };
            TestRange(testColumns);
        }

        [Test]
        public void IndexEncodingTest3Digits()
        {
            string[] testColumns = new string[] { "ZX", "ZY", "ZZ", "AAA", "AAB", "AAC" };
            TestRange(testColumns);
        }
        [Test]
        public void TestSingleIndex()
        {
            //const uint V = 283;
            const uint V = 2233;
            string v = ExcelColumnHelper.GetColumnIndexAsString(V);
            uint final = ExcelColumnHelper.GetColumnStringAsIndex(v);
        }


        [Test]
        public void IndexesToChars()
        {
            const uint START = 0;
            const uint MAX_END = 16384 * 20;
            const uint TEST_CAP = MAX_END - START;

            string[] testColumns = new string[TEST_CAP];
            for(uint i = START; i < TEST_CAP;i++)
            {
                testColumns[i - START] = ExcelColumnHelper.GetColumnIndexAsString(i);
            }

            for(uint i = START;i < TEST_CAP;i++)
            {
                string column = testColumns[i - START];
                uint result = ExcelColumnHelper.GetColumnStringAsIndex(column);
                Assert.IsTrue(result == i);
            }
        }
    }
}
