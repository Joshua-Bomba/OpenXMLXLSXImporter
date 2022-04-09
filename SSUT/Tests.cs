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
    public class Tests
    {
        public const string SHEET1 = "Sheet1";
        public const string SHEET2 = "Sheet2";
        public const string TEST_FILE = "tesxtcel.xlsx";

        public readonly static string[] EXPECTED_COLUMNS = new string[] { "Table Time", "Column1", "Column 2", "Column 3", "Data I want to Reference", "Maths", "Reference Math", "Another Sheet Reference", "Another Sheet Maths", null, null, null, null, null, "2", "0.26" };
        public readonly static string[] EXPECTED_ROWS = new string[] { null, null, null, "Reference Math", "433.26", "115.52", "104.53" };


        Stream stream = null;
        IExcelImporter importer = null;

        [SetUp]
        public void Setup()
        {
            stream = new MemoryStream();
            using(FileStream fs = File.OpenRead(TEST_FILE)) {
                fs.CopyTo(stream);
            }
            importer = new ExcelImporter(stream);
        }
        [TearDown]
        public void TearDown()
        {
            importer?.Dispose();
            stream?.Close();
            stream?.Dispose();
        }
        [Test]
        public void LoadTitleCellTestUsingRunner()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
            ICellData result = builder.Runner.LoadSingleCell(1, "A").GetAwaiter().GetResult();
            Assert.IsTrue(result.Content() == "A Header");
        }
        [Test]
        public void LoadTitleCellTestUsingBundler()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
            ISpreadSheetInstruction instruction = builder.Bundler.LoadSingleCell(1, "A");
            builder.Bundler.BundleRequeset(new List<ISpreadSheetInstruction> { instruction }).GetAwaiter().GetResult();
            ICellData result = instruction.GetResults().FirstAsync().GetAwaiter().GetResult();
            Assert.IsTrue(result.Content() == "A Header");
        }

        [Test]
        public void LoadTitleCellTestUsingBundlerResults()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
            ISpreadSheetInstruction instruction = builder.Bundler.LoadSingleCell(1, "A");
            IAsyncEnumerable<ICellData> results = builder.Bundler.GetBundledResults(new List<ISpreadSheetInstruction> { instruction});
            ICellData result = results.FirstAsync().GetAwaiter().GetResult();
            Assert.IsTrue(result.Content() == "A Header");
        }
        [Test]
        public void NonExistingSheet()
        {
            try
            {
                ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder("not a real sheet").GetAwaiter().GetResult();
                ICellData result = builder.Runner.LoadSingleCell(1, "A").GetAwaiter().GetResult();
                Assert.Fail();
            }
            catch (Exception ex)
            {

            }

        }
        [Test]
        public void TwoCells()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
            ISpreadSheetInstruction[] bundle = new ISpreadSheetInstruction[] {
                builder.Bundler.LoadSingleCell(2, "B"),
                builder.Bundler.LoadSingleCell(2, "A")
            };

            ICellData c = builder.Bundler.GetBundledResults(bundle).FirstAsync().GetAwaiter().GetResult();
             Assert.IsTrue(c.Content() == "Data in another cell");
        }
        [Test]
        public void ConcurrenyTest()
        {
            Task t = new Task(() =>
            {
                ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
                ICellData result = builder.Runner.LoadSingleCell(1, "A").GetAwaiter().GetResult();
                Assert.IsTrue(result.Content() == "A Header");
            });

            Task t2 = new Task(() =>
            {

                ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
                ICellData result = builder.Runner.LoadSingleCell(1, "A").GetAwaiter().GetResult();
                Assert.IsTrue(result.Content() == "A Header");
            });

            t.Start();
            t2.Start();

            t.Wait();
            t2.Wait();
        }
        [Test]
        public void RangeCellTest()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
            ValueTask<ICellData[]> columnRange = builder.Runner.LoadColumnRange(4, "A", "P").ToArrayAsync();
            ValueTask<ICellData[]> rowRange = builder.Runner.LoadRowRange("G", 1, 7).ToArrayAsync();
            CheckResultsAsync(columnRange, rowRange).GetAwaiter().GetResult();
        }

        public static async Task CheckResultsAsync(ValueTask<ICellData[]> columnRanges, ValueTask<ICellData[]> rowRanges)
        {
            CheckResults(await columnRanges, await rowRanges);
        }

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

        [Test]
        public void FullRangeCellTest()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
            ValueTask<ICellData[]> columnRange = builder.Runner.LoadFullColumnRange(4).ToArrayAsync();
            ValueTask<ICellData[]> rowRange = builder.Runner.LoadFullRowRange("G").ToArrayAsync();
            CheckResultsAsync(columnRange, rowRange).GetAwaiter().GetResult();
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
        [Test]
        public void FullRangeCellTestLoadTheSameDataAgain()
        {
            Parallel.For(0, 100000, new ParallelOptions { }, (i) =>
            {
                    try
                    {
                        ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
                        ValueTask<ICellData[]> columnRange = builder.Runner.LoadFullColumnRange(4).ToArrayAsync();
                        ValueTask<ICellData[]> rowRange = builder.Runner.LoadFullRowRange("G").ToArrayAsync();
                        CheckResultsAsync(columnRange, rowRange).GetAwaiter().GetResult();
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
                using (FileStream fs = File.OpenRead(TEST_FILE))
                {
                    fs.CopyTo(baseStream);
                }
                data = baseStream.ToArray();
            }

            Parallel.For(0, 100000, new ParallelOptions { MaxDegreeOfParallelism = 2 }, (i) =>
            {
                using (MemoryStream ms = new MemoryStream(data, false))
                {
                    using (IExcelImporter importer = new ExcelImporter(ms))
                    {
                        try
                        {
                            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
                            ValueTask<ICellData[]> columnRange = builder.Runner.LoadFullColumnRange(4).ToArrayAsync();
                            ValueTask<ICellData[]> rowRange = builder.Runner.LoadFullRowRange("G").ToArrayAsync();
                            CheckResultsAsync(columnRange, rowRange).GetAwaiter().GetResult();
                        }
                        catch
                        {

                        }

                    }
                }
            });

        }



        //[Test]
        //public void CloseConnectionToEarly()
        //{
        //    ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1).GetAwaiter().GetResult();
        //}
    }
}