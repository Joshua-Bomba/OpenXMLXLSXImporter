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
    public class LastColumn : ICellProcessingTask
    {
        public uint _row;
        private ICellIndex _result;
        private AsyncManualResetEvent _mre;
        private Exception _fail;
        public LastColumn(uint row)
        {
            _mre = new AsyncManualResetEvent(false);
            _row = row;
        }

        public void Failure(Exception e)
        {
            _fail = e;
            _mre.Set();
        }

        public async Task<ICellIndex> GetIndex()
        {
            await _mre.WaitAsync();
            if(_fail != null)
            {
                throw _fail;
            }
            return _result;
        }

        public void Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
        {
            _result = index;
            _mre.Set();
        }
    }
}
