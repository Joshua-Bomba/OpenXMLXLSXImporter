using System;
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
using OpenXMLXLSXImporter.Builders.Managers;
using OpenXMLXLSXImporter.CellData;
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
        
       

        private Dictionary<string, ISpreadSheet> _loadedSheets;

        public ExcelImporter(Stream stream)
        {
            _streamSheetFile = new XlsxDocumentFile(stream);
            _loadedSheets = new Dictionary<string, ISpreadSheet>();
        }



        public async Task<ISpreadSheet> GetSheet(string sheet)
        {
            IXlsxDocumentFile df = await _streamSheetFile.GetLoadedFile();
            if (df.ContainsSheet(sheet))
            {
                if (!_loadedSheets.ContainsKey(sheet))
                {
                    //this is the first time we use this sheet
                    _loadedSheets[sheet] = new SpreadSheet(new XlsxSheetFile(df, sheet));
                }
                return _loadedSheets[sheet];
            }
            return null;
        }

        public void Dispose()
        {
            foreach(ISpreadSheet iss in _loadedSheets.Values)
            {
                iss?.Dispose();
            }
        }


        //public Task Process(ISpreadSheet sheet)
        //{
        //    return Task.Run(async() =>
        //    {
        //        try
        //        {
        //            SpreadSheetInstructionBuilder ssib = new SpreadSheetInstructionBuilder();
        //            Task<IXlsxSheetFilePromise> gt = _streamSheetFile.LoadSpreadSheetData(sheet);
        //            sheet.LoadConfig(ssib);
        //            IXlsxSheetFilePromise g = await gt;
        //            dequeManager = new SpreadSheetDequeManager(this);
        //            ssim = new SpreadSheetInstructionManager(dequeManager);
        //            dequeManager.StartRequestProcessor(g);
        //            await ssib.ProcessInstructions(ssim);
        //            Task r = ssib.ProcessResults();
        //            await sheet.ResultsProcessed(ssib);
        //            await r;
        //        }
        //        catch(Exception ex)
        //        {
        //            ExceptionDispatchInfo.Capture(ex).Throw();
        //        }

        //    });

        //}



        //public async IAsyncEnumerable<IEnumerable<Task<ICellData>>> ProcessAndGetAsyncCollection(string sheetName, Action<ISpreadSheet> builder)
        //{
        //    SpreadSheet dssp = new SpreadSheet(sheetName,builder);
        //    await this.Process(dssp);
        //    IAsyncEnumerable<IEnumerable<Task<ICellData>>> ret = dssp.GetResults();
        //    IAsyncEnumerator<IEnumerable<Task<ICellData>>> f = ret.GetAsyncEnumerator();
        //    while(await f.MoveNextAsync())
        //    {
        //        yield return f.Current;
        //    }
        //}

        //public async Task<List<List<ICellData>>> ProcessAndGetListAsync(string sheetName, Action<ISpreadSheet> builder)
        //{
        //    SpreadSheet dssp = new SpreadSheet(sheetName, builder);
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
