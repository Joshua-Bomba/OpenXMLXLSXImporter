using Nito.AsyncEx;
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
            _row = row;
        }

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
