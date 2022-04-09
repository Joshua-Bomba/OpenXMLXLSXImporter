using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.FileAccess;
using OpenXMLXLSXImporter.Indexers;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public class DeferredCell : IFutureCell, ICellIndex
    {
        private Cell _deferredCell;
        private AsyncLock _lock;
        private DeferredCellExecution _deferredCellExecution;
        public DeferredCell(uint cellRowIndex, string cellColumnIndex, Cell cell)
        {
            CellRowIndex = cellRowIndex;
            CellColumnIndex = cellColumnIndex;
            _lock = new AsyncLock();
            _deferredCell = cell;
            _deferredCellExecution = null;
        }

        public ISpreadSheetInstructionManager InstructionManager { get; set; }

        public IFutureUpdate Updater { get; set; }

        private class DeferredCellExecution : ICellProcessingTask, IFutureCell
        {
            private DeferredCell _deferredCell;
            private ICellData _result;
            private AsyncManualResetEvent _mre;
            private IFutureUpdate _updater;
            public DeferredCellExecution(DeferredCell deferredCell, IFutureUpdate updater)
            {
                _mre = new AsyncManualResetEvent(false);
                _deferredCell = deferredCell;
                _updater = updater;
            }

            public async Task<ICellData> GetData()
            {
                await _mre.WaitAsync();
                return _result;
            }

            public void Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
            {
                _result = file.ProcessedCell(_deferredCell._deferredCell, _deferredCell);
                _updater.Update(_result);
                _mre.Set();
            }
        }

        public string CellColumnIndex { get; set; }
        public uint CellRowIndex { get; set; }

        public async Task<ICellData> GetData()
        {
            if(_deferredCellExecution == null)
            {
                using(await _lock.LockAsync())
                {
                    if (_deferredCellExecution == null)
                    {
                        _deferredCellExecution = new DeferredCellExecution(this,Updater);
                        await InstructionManager.Queue.QueueNonIndexedCell(_deferredCellExecution);
                    }
                }
            }
            return await _deferredCellExecution.GetData();
        }
    }
}
