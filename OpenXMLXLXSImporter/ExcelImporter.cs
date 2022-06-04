using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.CellParsing;
using OpenXMLXLSXImporter.FileAccess;
using OpenXMLXLSXImporter.Processing;

namespace OpenXMLXLSXImporter
{

    /// <summary>
    /// This class is for managing the import pass in a Stream to your excel file
    /// Implemented the ISheetProperties interface for each sheet you want to import
    /// add the ISheetProperties through the Add Method
    /// This will make it load the sheet in
    /// </summary>
    public class ExcelImporter : IExcelImporter
    {
        private XlsxDocumentFile _streamSheetFile;

        private ConcurrentDictionary<string, SpreadSheetInstructionManager> _instructionBuilders;
        private ICellParser _cellParser;
        public ExcelImporter(Stream stream, ICellParser parser)
        {
            _cellParser = parser;
            _streamSheetFile = new XlsxDocumentFile(stream);
            _instructionBuilders = new ConcurrentDictionary<string, SpreadSheetInstructionManager>();
        }

        public ExcelImporter(Stream stream) : this(stream, new DefaultCellParser()) { }


        public async Task<ISpreadSheetInstructionBuilder> GetSheetBuilder(string sheetName)
            =>new SpreadSheetInstructionBuilder(_instructionBuilders.GetOrAdd(sheetName, x => new SpreadSheetInstructionManager(_streamSheetFile.LoadSpreadSheetData(x, _cellParser))));

        public void Dispose()
        {
            ConcurrentDictionary<string, SpreadSheetInstructionManager> end = _instructionBuilders;
            _instructionBuilders = null;
            foreach (KeyValuePair<string, SpreadSheetInstructionManager> ssib in end)
            {
                ssib.Value.Dispose();
            }
            _streamSheetFile.Dispose();
        }

        //public async IAsyncEnumerable<IEnumerable<Task<ICellData>>> ProcessAndGetAsyncCollection(string sheetName, Action<ISpreadSheetInstructionBuilderManagerInstructionBuilder> builder)
        //{
        //    SpreadSheetInstructionBuilderManager dssp = new SpreadSheetInstructionBuilderManager(sheetName,builder);
        //    await this.Process(dssp);
        //    IEnumerable<IAsyncEnumerable<ICellData>> ret = dssp.GetResults();
        //    foreach(IAsyncEnumerable<ICellData> d in ret)
        //    {

        //    }
        //    IAsyncEnumerator<IEnumerable<Task<ICellData>>> f = ret.GetAsyncEnumerator();
        //    while(await f.MoveNextAsync())
        //    {
        //        yield return f.Current;
        //    }
        //}

        //public async Task<List<List<ICellData>>> ProcessAndGetListAsync(string sheetName, Action<ISpreadSheetInstructionBuilderManagerInstructionBuilder> builder)
        //{
        //    SpreadSheetInstructionBuilderManager dssp = new SpreadSheetInstructionBuilderManager(sheetName, builder);
        //    Task pt =  this.Process(dssp);
        //    List<List<ICellData>> output = new List<List<ICellData>>();
        //    await pt;
        //    IAsyncEnumerable<IEnumerable<Task<ICellData>>> ret = dssp.GetResults();
        //    IAsyncEnumerator<IEnumerable<Task<ICellData>>> f = ret.GetAsyncEnumerator();

        //    List<ICellData> currentList;

        //    while(await f.MoveNextAsync())
        //    {
        //        currentList = new List<ICellData>();
        //        foreach (Task<ICellData> l2 in f.Current)
        //        {
        //            currentList.Add(await l2);
        //        }
        //        output.Add(currentList);
        //    }
        //    return output;
        //}
    }
}
