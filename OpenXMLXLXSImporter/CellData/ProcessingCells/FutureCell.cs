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
        private IFutureUpdate _updater;
        private AsyncManualResetEvent _mre;
        public FutureCell(uint cellRowIndex, string cellColumnIndex, IFutureUpdate updater)
        {
            CellRowIndex = cellRowIndex;
            CellColumnIndex = cellColumnIndex;
            _mre = new AsyncManualResetEvent(false);
            _updater = updater;
        }
        public string CellColumnIndex { get; set; }

        public uint CellRowIndex { get; set; }

        public async Task<ICellData> GetData()
        {
            await _mre.WaitAsync();
            return _result;
        }

        public void Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
        {
            try
            {
                _result = file.ProcessedCell(cellElement, index);
                _updater?.Update(_result);
            }
            finally
            {
                _mre.Set();
            }

        }
    }
}
