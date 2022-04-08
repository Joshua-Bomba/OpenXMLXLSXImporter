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
    public class LastRow : ICellProcessingTask
    {
        private ICellIndex _result;
        private AsyncManualResetEvent _mre;
        public LastRow()
        {
            _mre = new AsyncManualResetEvent(false);
        }

        public async Task<ICellIndex> GetIndex()
        {
            await _mre.WaitAsync();
            return _result;
        }
        public void Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
        {
            _result = index;
            _mre.Set();
        }
    }
}
