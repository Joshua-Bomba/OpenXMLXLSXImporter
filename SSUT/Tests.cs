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
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1);
            ICellData result = builder.Runner.LoadSingleCell(1, "A").GetAwaiter().GetResult();
            Assert.IsTrue(result.Content() == "A Header");
        }
        [Test]
        public void LoadTitleCellTestUsingBundler()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1);
            ISpreadSheetInstruction instruction = builder.Bundler.LoadSingleCell(1, "A");
            builder.Bundler.BundleRequeset(new List<ISpreadSheetInstruction> { instruction }).GetAwaiter().GetResult();
            ICellData result = instruction.GetResults().FirstAsync().GetAwaiter().GetResult();
            Assert.IsTrue(result.Content() == "A Header");
        }

        [Test]
        public void LoadTitleCellTestUsingBundlerResults()
        {
            ISpreadSheetInstructionBuilder builder = importer.GetSheetBuilder(SHEET1);
            ISpreadSheetInstruction instruction = builder.Bundler.LoadSingleCell(1, "A");
            IAsyncEnumerable<ICellData> results = builder.Bundler.GetBundledResults(new List<ISpreadSheetInstruction> { instruction});
            ICellData result = results.FirstAsync().GetAwaiter().GetResult();
            Assert.IsTrue(result.Content() == "A Header");
        }

        //[Test]
        //public void LoadTitleCellWithoutImplementAInterface()
        //{
        //    ICellData c = null;
        //    var f = async () =>
        //    {

        //        // IAsyncEnumerable<IEnumerable<Task<ICellData>>> ret = importer.ProcessAndGetAsyncCollection(SHEET1, x => x.LoadSingleCell(1, "A"));
        //        //c = await (await ret.FirstAsync()).First();

        //    };

        //    f().GetAwaiter().GetResult();
        //    Assert.IsTrue(c.Content() == "A Header");
            
        //}
        //[Test]
        //public void LoadTitleCellWithoutImplementAInterfaceList()
        //{
        //    ICellData c;
        //    //List<List<ICellData>> ret =  importer.ProcessAndGetListAsync(SHEET1, x => x.LoadSingleCell(1, "A")).GetAwaiter().GetResult();
        //    c = ret.First().First();
        //    Assert.IsTrue(c.Content() == "A Header");
        //}
        //[Test]
        //public void SecondCell()
        //{
        //    ICellData c;
        //    //List<List<ICellData>> ret = importer.ProcessAndGetListAsync(SHEET1, x => x.LoadSingleCell(2, "B").LoadSingleCell(2,"A")).GetAwaiter().GetResult();
        //    //c = ret.First().First();
        //   // Assert.IsTrue(c.Content() == "Data in another cell");
        //}
        //private class OnlySecondCell : ISpreadSheetInstructionBuilderManager
        //{
        //    private ISpreadSheetInstructionKey secondCell;
        //    public OnlySecondCell()
        //    {

        //    }
        //    public string Sheet => SHEET1;

        //    public void LoadConfig(ISpreadSheetInstructionBuilder builder)
        //    {
        //        secondCell = builder.LoadSingleCell(2, "B");
        //        //we will wait till we select the second cell before we select the first one. it should queue it in the defered area given enought time
                
        //    }

        //    public async Task ResultsProcessed(ISpreadSheetQueryResults query)
        //    {
        //        ICellData cell = await query.GetProcessedResults(secondCell).FirstOrDefaultAsync();

        //    }
        //}

        //private class PostFirstCell : ISpreadSheetInstructionBuilderManager
        //{
        //    private ISpreadSheetInstructionKey firstCell;
        //    public string Sheet => SHEET1;

        //    public void LoadConfig(ISpreadSheetInstructionBuilder builder)
        //    {
        //        firstCell = builder.LoadSingleCell(2, "A");
        //    }

        //    public async Task ResultsProcessed(ISpreadSheetQueryResults query)
        //    {
        //        ICellData d = await query.GetProcessedResults(firstCell).FirstOrDefaultAsync();
        //    }
        //}

        //[Test]
        //public void OnlySecondCellTest()
        //{
        //    OnlySecondCell osc = new OnlySecondCell();
        //    importer.Process(osc).GetAwaiter().GetResult();
        //    System.Threading.Thread.Sleep(10000);
        //    importer.Process(osc).GetAwaiter().GetResult();
        //}

        //private class RangeCellsTest : ISpreadSheetInstructionBuilderManager
        //{
        //    public string Sheet => SHEET1;
        //    ISpreadSheetInstructionKey _columnRangek;
        //    ISpreadSheetInstructionKey _rowRangek;

        //    public List<ICellData> columnRanges;
        //    public List<ICellData> rowRanges;

        //    public void LoadConfig(ISpreadSheetInstructionBuilder builder)
        //    {
        //        _columnRangek = builder.LoadColumnRange(4, "A", "P");
        //        _rowRangek = builder.LoadRowRange("G", 1, 7);
        //    }

        //    public async Task ResultsProcessed(ISpreadSheetQueryResults query)
        //    {
        //        IAsyncEnumerable<ICellData> columnRange = query.GetProcessedResults(_columnRangek);
        //        IAsyncEnumerable<ICellData> rowRange = query.GetProcessedResults(_rowRangek);

        //        columnRanges = new List<ICellData>();
        //        rowRanges = new List<ICellData>();

        //        IAsyncEnumerator<ICellData> columnEnumerator = columnRange.GetAsyncEnumerator();

        //        while(await columnEnumerator.MoveNextAsync())
        //        {
        //            ICellData d = columnEnumerator.Current;
        //            columnRanges.Add(d);
        //        }

        //        IAsyncEnumerator<ICellData> RowEnumerator = rowRange.GetAsyncEnumerator();

        //        while(await RowEnumerator.MoveNextAsync())
        //        {
        //            rowRanges.Add(RowEnumerator.Current);
        //        }

        //    }
        //}
        //[Test]
        //public void RangeCellTest()
        //{
        //    RangeCellsTest rct = new RangeCellsTest();
        //    importer.Process(rct).GetAwaiter().GetResult();
        //    CheckResults(rct.columnRanges, rct.rowRanges);
        //}

        //public static void CheckResults(List<ICellData> columnRanges, List<ICellData> rowRanges)
        //{
        //    Assert.IsTrue(EXPECTED_COLUMNS.Length == columnRanges.Count);
        //    Assert.IsTrue(EXPECTED_ROWS.Length == rowRanges.Count);
        //    for (int i = 0; i < EXPECTED_COLUMNS.Length; i++)
        //    {
        //        Assert.AreEqual(EXPECTED_COLUMNS[i], columnRanges[i].Content());
        //    }

        //    for (int i = 0; i < EXPECTED_ROWS.Length; i++)
        //    {
        //        Assert.AreEqual(EXPECTED_ROWS[i], rowRanges[i].Content());
        //    }
        //}

        //private class FullRangeCellsTest: ISpreadSheetInstructionBuilderManager
        //{
        //    public string Sheet => SHEET1;
        //    ISpreadSheetInstructionKey _columnRange;
        //    ISpreadSheetInstructionKey _rowRange;

        //    public List<ICellData> columnRanges;
        //    public List<ICellData> rowRanges;
        //    public FullRangeCellsTest()
        //    {
        //        columnRanges = null;
        //        rowRanges = null;
        //    }

        //    public void LoadConfig(ISpreadSheetInstructionBuilder builder)
        //    {
        //        _columnRange = builder.LoadFullColumnRange(4);
        //        _rowRange = builder.LoadFullRowRange("G");
        //    }

        //    public void CompareResults()
        //    {

        //    }

        //    public async Task ResultsProcessed(ISpreadSheetQueryResults query)
        //    {
        //        IAsyncEnumerable<ICellData> columnRange = query.GetProcessedResults(_columnRange);
        //        IAsyncEnumerable<ICellData> rowRange = query.GetProcessedResults(_rowRange);
        //        columnRanges = new List<ICellData>();
        //        rowRanges = new List<ICellData>();

        //        IAsyncEnumerator<ICellData> columnEnumerator = columnRange.GetAsyncEnumerator();

        //        while (await columnEnumerator.MoveNextAsync())
        //        {
        //            ICellData d = columnEnumerator.Current;
        //            columnRanges.Add(d);
        //        }

        //        IAsyncEnumerator<ICellData> RowEnumerator = rowRange.GetAsyncEnumerator();

        //        while (await RowEnumerator.MoveNextAsync())
        //        {
        //            rowRanges.Add(RowEnumerator.Current);
        //        }
        //    }
        //}

        //[Test]
        //public void FullRangeCellTest()
        //{
        //    FullRangeCellsTest rc = new FullRangeCellsTest();
        //    importer.Process(rc).GetAwaiter().GetResult();
        //    CheckResults(rc.columnRanges, rc.rowRanges);
        //}

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


        //[Test]
        //public void FullRangeCellTestBurnInTest()
        //{
        //    byte[] data;
        //    using(MemoryStream baseStream = new MemoryStream())
        //    {
        //        using (FileStream fs = File.OpenRead(TEST_FILE))
        //        {
        //            fs.CopyTo(baseStream);
        //        }
        //        data = baseStream.ToArray();
        //    }

        //    Parallel.For(0, 100000,new ParallelOptions { MaxDegreeOfParallelism = 2 } , (i) =>
        //    {
        //        using (MemoryStream ms = new MemoryStream(data,false))
        //        {
        //            using (IExcelImporter importer = new ExcelImporter(ms))
        //            {
        //                try
        //                {
        //                    FullRangeCellsTest rc = new FullRangeCellsTest();
        //                    importer.Process(rc).GetAwaiter().GetResult();
        //                    CheckResults(rc.columnRanges, rc.rowRanges);
        //                }
        //                catch
        //                {

        //                }

        //            }
        //        }
        //    });

        //}
    }
}