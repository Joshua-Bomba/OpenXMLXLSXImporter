using Nito.AsyncEx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public class LastRow : ICellProcessingTask, IFutureCell
    {
        public string _column;
        private ICellData _result;
        private AsyncManualResetEvent _mre;
        public LastRow(string column)
        {
            _column = column;
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
