using NUnit.Framework;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSUT
{
    public static class Global
    {
        public const string SHEET1 = "Sheet1";
        public const string SHEET2 = "Sheet2";
        public const string TEST_FILE = "tesxtcel.xlsx";

        public readonly static string[] EXPECTED_COLUMNS = new string[] { "Table Time", "Column1", "Column 2", "Column 3", "Data I want to Reference", "Maths", "Reference Math", "Another Sheet Reference", "Another Sheet Maths", null, null, null, null, null, "2", "0.26" };
        public readonly static string[] EXPECTED_ROWS = new string[] { null, null, null, "Reference Math", "433.26", "115.52", "104.53" };

        public static void CheckResults(ICellData[] columnRanges, ICellData[] rowRanges)
        {
            Assert.IsTrue(EXPECTED_COLUMNS.Length == columnRanges.Length);
            Assert.IsTrue(EXPECTED_ROWS.Length == rowRanges.Length);
            for (int i = 0; i < EXPECTED_COLUMNS.Length; i++)
            {
                Assert.AreEqual(EXPECTED_COLUMNS[i], columnRanges[i].Content());
            }

            for (int i = 0; i < EXPECTED_ROWS.Length; i++)
            {
                Assert.AreEqual(EXPECTED_ROWS[i], rowRanges[i].Content());
            }
        }

        public static async Task CheckResultsAsync(ValueTask<ICellData[]> columnRanges, ValueTask<ICellData[]> rowRanges)
        {
            CheckResults(await columnRanges, await rowRanges);
        }

        //internal class testThing<T>
        //{
        //    public HashSet<T> set;
        //    private object l;

        //    public testThing()
        //    {
        //        set = new HashSet<T>();
        //        l = new object();
        //    }
        //    public void LogEnque(T c)
        //    {
        //        Task.Run(() =>
        //        {
        //            lock (l)
        //            {
        //                set.Add(c);
        //            }
        //        });
        //    }

        //    public void LogDeque(T c)
        //    {
        //        Task.Run(() =>
        //        {
        //            lock (l)
        //            {
        //                if (set.Contains(c))
        //                {
        //                    set.Remove(c);
        //                }
        //            }
        //        });
        //    }
        //}
    }
}
