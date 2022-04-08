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
        private IFutureUpdate _updater;
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

        private class DeferredCellExecution : ICellProcessingTask, IFutureCell
        {
            private DeferredCell _deferredCell;
            private ICellData _result;
            private AsyncManualResetEvent _mre;
            private IFutureUpdate _updater;
            public DeferredCellExecution(DeferredCell deferredCell)
            {
                _mre = new AsyncManualResetEvent(false);
                _deferredCell = deferredCell;
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

            public void Updateder(IFutureUpdate cellUpdater)
            {
                _updater = cellUpdater;
            }
        }

        public ISpreadSheetInstructionManager InstructionManager { get; set; }

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
                        _deferredCellExecution = new DeferredCellExecution(this);
                        _deferredCellExecution.Updateder(_updater);
                        await InstructionManager.Queue.QueueNonIndexedCell(_deferredCellExecution);
                    }
                }
            }
            return await _deferredCellExecution.GetData();
        }
        public void Updateder(IFutureUpdate cellUpdater)
        {
            _updater = cellUpdater;
        }
    }
}
