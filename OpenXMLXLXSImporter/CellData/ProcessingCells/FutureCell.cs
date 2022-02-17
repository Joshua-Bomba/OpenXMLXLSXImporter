using Nito.AsyncEx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public class FutureCell : IFutureCell, ICellProcessingTask
    {
        private  ICellData _result;
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
            return _result;
        }

        public void Resolve(ICellData data)
        {
            _result = data;
            _mre.Set();
        }
    }
}
