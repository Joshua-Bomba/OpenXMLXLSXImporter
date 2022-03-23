using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLSXImporter.FileAccess;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public class FutureCell : IFutureCell, ICellProcessingTask, ICellIndex
    {
        private  ICellData _result;
        private IIndexer _indexer;
        private AsyncManualResetEvent _mre;
        public FutureCell(uint cellRowIndex, string cellColumnIndex)
        {
            CellRowIndex = cellRowIndex;
            CellColumnIndex = cellColumnIndex;
            _mre = new AsyncManualResetEvent(false);
        }
        public string CellColumnIndex { get; set; }

        public uint CellRowIndex { get; set; }

        public async Task<ICellData> GetData()
        {
            await _mre.WaitAsync();
            await _indexer.SetCell(_result);
            return _result;
        }

        public void Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
        {
            _result = file.ProcessedCell(cellElement, index);
            _mre.Set();
        }
        public void SetIndexer(IIndexer indexer)
        {
            _indexer = indexer;
        }
    }
}
