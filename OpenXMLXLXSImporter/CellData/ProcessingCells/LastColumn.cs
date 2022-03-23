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
    public class LastColumn : ICellProcessingTask, IFutureCell
    {
        public uint _row;
        private ICellData _result;
        private AsyncManualResetEvent _mre;
        public LastColumn(uint row)
        {
            _mre = new AsyncManualResetEvent(false);
            _row = row;
        }

        public async Task<ICellData> GetData()
        {
            await _mre.WaitAsync();
            return _result;
        }

        public void Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
        {
            _result = file.ProcessedCell(cellElement, index);
            _mre.Set();
        }

        public void SetIndexer(IIndexer indexer) { }
    }
}
