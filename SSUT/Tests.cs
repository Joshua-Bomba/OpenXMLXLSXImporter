using NUnit.Framework;
using OpenXMLXLSXImporter;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.Builders.Managers;
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
        Stream stream = null;
        IExcelImporter importer = null;

        [SetUp]
        public void Setup()
        {
            stream = File.OpenRead(TEST_FILE);
            importer = new ExcelImporter(stream);
        }
        [TearDown]
        public void TearDown()
        {
            importer?.Dispose();
            stream?.Close();
            stream?.Dispose();
        }

        private class TitleColumnTest : ISpreadSheetInstructionBuilderManager
        {
            private ISpreadSheetInstructionKey title;

            public string Sheet => SHEET1;

            public void LoadConfig(ISpreadSheetInstructionBuilder builder)
            {
                title = builder.LoadSingleCell(1, "A");
            }

            public async Task ResultsProcessed(ISpreadSheetQueryResults query)
            {
                ICellData d = await query.GetProcessedResults(title).FirstOrDefaultAsync();

                Assert.IsTrue(d.Content() == "A Header");
            }
        }

        [Test]
        public void LoadTitleCellTest()
        {
            TitleColumnTest titleColumnTest = new TitleColumnTest();
            importer.Process(titleColumnTest).GetAwaiter().GetResult();
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
        private class OnlySecondCell : ISpreadSheetInstructionBuilderManager
        {
            private ISpreadSheetInstructionKey secondCell;
            public OnlySecondCell()
            {

            }
            public string Sheet => SHEET1;

            public void LoadConfig(ISpreadSheetInstructionBuilder builder)
            {
                secondCell = builder.LoadSingleCell(2, "B");
                //we will wait till we select the second cell before we select the first one. it should queue it in the defered area given enought time
                
            }

            public async Task ResultsProcessed(ISpreadSheetQueryResults query)
            {
                ICellData cell = await query.GetProcessedResults(secondCell).FirstOrDefaultAsync();

            }
        }

        private class PostFirstCell : ISpreadSheetInstructionBuilderManager
        {
            private ISpreadSheetInstructionKey firstCell;
            public string Sheet => SHEET1;

            public void LoadConfig(ISpreadSheetInstructionBuilder builder)
            {
                firstCell = builder.LoadSingleCell(2, "A");
            }

            public async Task ResultsProcessed(ISpreadSheetQueryResults query)
            {
                ICellData d = await query.GetProcessedResults(firstCell).FirstOrDefaultAsync();
            }
        }

        [Test]
        public void OnlySecondCellTest()
        {
            OnlySecondCell osc = new OnlySecondCell();
            importer.Process(osc).GetAwaiter().GetResult();
            System.Threading.Thread.Sleep(10000);
            importer.Process(osc).GetAwaiter().GetResult();
        }

        private class RangeCellsTest : ISpreadSheetInstructionBuilderManager
        {
            public string Sheet => SHEET1;
            ISpreadSheetInstructionKey _columnRange;
            ISpreadSheetInstructionKey _rowRange;

            public void LoadConfig(ISpreadSheetInstructionBuilder builder)
            {
                _columnRange = builder.LoadColumnRange(4, "A", "I");
                _rowRange = builder.LoadRowRange("G", 5, 7);
            }

            public async Task ResultsProcessed(ISpreadSheetQueryResults query)
            {
                IAsyncEnumerable<ICellData> columnRange = query.GetProcessedResults(_columnRange);
                IAsyncEnumerable<ICellData> rowRange = query.GetProcessedResults(_rowRange);

                List<string> columnRanges = new List<string>();
                List<string> rowRanges = new List<string>();

                IAsyncEnumerator<ICellData> columnEnumerator = columnRange.GetAsyncEnumerator();

                while(await columnEnumerator.MoveNextAsync())
                {
                    ICellData d = columnEnumerator.Current;
                    columnRanges.Add(d.Content());
                }

                IAsyncEnumerator<ICellData> RowEnumerator = rowRange.GetAsyncEnumerator();

                while(await RowEnumerator.MoveNextAsync())
                {
                    rowRanges.Add(RowEnumerator.Current.Content());
                }

            }
        }
        [Test]
        public void RangeCellTest()
        {
            RangeCellsTest rct = new RangeCellsTest();
            importer.Process(rct).GetAwaiter().GetResult();
        }

        private class FullRangeCellsTest: ISpreadSheetInstructionBuilderManager
        {
            public string Sheet => SHEET1;

            public void LoadConfig(ISpreadSheetInstructionBuilder builder)
            {
                throw new NotImplementedException();
            }

            public Task ResultsProcessed(ISpreadSheetQueryResults query)
            {
                throw new NotImplementedException();
            }
        }
        [Test]
        public void FullRangeCellTest()
        {
            FullRangeCellsTest rc = new FullRangeCellsTest();
            importer.Process(rc).GetAwaiter().GetResult();
        }
    }
}