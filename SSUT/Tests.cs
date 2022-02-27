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

        //private class TitleColumnTest : ISpreadSheet
        //{
        //    private ISpreadSheetInstructionKey title;

        //    public string Sheet => SHEET1;

        //    public void LoadConfig(ISpreadSheetInstructionBuilder builder)
        //    {
        //        title = builder.LoadSingleCell(1, "A");
        //    }

        //    public async Task ResultsProcessed(ISpreadSheetQueryResults query)
        //    {
        //        ICellData d = await (await query.GetResults(title)).FirstOrDefault();

        //        Assert.IsTrue(d.Content() == "A Header");
        //    }
        //}

        [Test]
        public void LoadTitleCellTest()
        {
            ICellData c = null;
            var f = async () =>
            {

                ISpreadSheet sheet = await importer.GetSheet(SHEET1);
                c = await sheet.LoadSingleCell(1, "A");
            };

            f().GetAwaiter().GetResult();
            Assert.IsTrue(c.Content() == "A Header");
        }
        //[Test]
        //public void LoadTitleCellWithoutImplementAInterfaceList()
        //{
        //    ICellData c;
        //    List<List<ICellData>> ret =  importer.ProcessAndGetListAsync(SHEET1, x => x.LoadSingleCell(1, "A")).GetAwaiter().GetResult();
        //    c = ret.First().First();
        //    Assert.IsTrue(c.Content() == "A Header");
        //}
        //[Test]
        //public void SecondCell()
        //{
        //    ICellData c;
        //    List<List<ICellData>> ret = importer.ProcessAndGetListAsync(SHEET1, x => x.LoadSingleCell(2, "B").LoadSingleCell(2,"A")).GetAwaiter().GetResult();
        //    c = ret.First().First();
        //    Assert.IsTrue(c.Content() == "Data in another cell");
        //}

        //[Test]
        //public void OnlySecondCell()
        //{
        //    ICellData c;
        //    List<List<ICellData>> ret = importer.ProcessAndGetListAsync(SHEET1, x => x.LoadSingleCell(2, "B")).GetAwaiter().GetResult();
        //    c = ret.First().First();
        //    //we will wait till we select the second cell before we select the first one. it should queue it in the defered area given enought time
        //    System.Threading.Thread.Sleep(10000);
        //    //...later
        //    //it should be in a deferred state and we need to get the value here
        //    List<List<ICellData>> ret2 = importer.ProcessAndGetListAsync(SHEET1, x => x.LoadSingleCell(2, "A")).GetAwaiter().GetResult();
        //    Assert.IsTrue(c.Content() == "Data in another cell");
        //}
    }
}