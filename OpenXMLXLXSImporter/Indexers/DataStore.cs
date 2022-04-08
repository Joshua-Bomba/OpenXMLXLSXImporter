using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public class DataStore : IDataStore
    {
        protected SpreadSheetInstructionManager _instructionManager;

        protected RowIndexer _rowIndexer;
        private AsyncLock _accessorLock = new AsyncLock();
        public DataStore(SpreadSheetInstructionManager instructionManager)
        {
            _rowIndexer = new RowIndexer();
            _instructionManager = instructionManager;
        }
        public async Task<LastColumn> GetLastColumn(uint rowIndex)
        {
            if(!_rowIndexer.ContainsKey(rowIndex)||_rowIndexer[rowIndex].LastColumn == null)
            {
                using (await _accessorLock.LockAsync())
                {
                    if(!_rowIndexer.ContainsKey(rowIndex))
                    {
                        _rowIndexer[rowIndex] = new ColumnIndexer();
                    }
                    if (_rowIndexer[rowIndex].LastColumn == null)
                    {
                        _rowIndexer[rowIndex].LastColumn = new LastColumn(rowIndex);
                        await this.QueueNonIndexedCell(_rowIndexer[rowIndex].LastColumn);
                    }
                }
            }
            return _rowIndexer[rowIndex].LastColumn;
        }

        public async Task<LastRow> GetLastRow()
        {
            if(_rowIndexer.LastRow == null)
            {
                using (await _accessorLock.LockAsync())
                {
                    if(_rowIndexer.LastRow == null)
                    {
                        _rowIndexer.LastRow = new LastRow();
                        await this.QueueNonIndexedCell(_rowIndexer.LastRow);
                    }
                }
            }
            return _rowIndexer.LastRow;
        }

        public async Task<ICellIndex> GetCell(uint rowIndex, string cellIndex, Func<ICellProcessingTask> newCell = null)
        {
            ICellIndex result = _rowIndexer.Get(rowIndex, cellIndex);
            if (result == null&&newCell != null)
            {
                using (await this._instructionManager.Queue.Mutex.LockAsync())
                {
                    ICellIndex r = _rowIndexer.Get(rowIndex, cellIndex);
                    if (r == null)
                    {
                        ICellProcessingTask t = newCell();
                        if (t != null)
                        {
                            if (t is IFutureCell fc)
                            {
                                fc.SetDataStore(this);
                            }
                            this._instructionManager.Queue.Enque(t);
                            if (t is ICellIndex index)
                            {
                                r = index;
                                _rowIndexer.Set(index);
                            }
                        }
                       
                    }
                    return r;
                }
            }
            return result;
        }
        public virtual async Task ProcessInstruction(ISpreadSheetInstruction instruction)
        {
            using (await _accessorLock.LockAsync())
            {
                await instruction.EnqueCell(this);
            }
        }

        public async Task SetCell(ICellIndex index)
        {
            using(await _accessorLock.LockAsync())
            {
                if(index is IFutureCell fc)
                {
                    fc.SetDataStore(this);
                }
                _rowIndexer.Set(index);
            }
        }

        public async Task QueueNonIndexedCell(ICellProcessingTask t)
        {
            using (await _instructionManager.Queue.Mutex.LockAsync())
            {
                _instructionManager.Queue.Enque(t);
            }
        }

        public async Task SetCells(IEnumerable<ICellIndex> cells)
        {
            using (await _accessorLock.LockAsync())
            {
                foreach(ICellIndex cell in cells)
                {
                    if (cell is IFutureCell fc)
                    {
                        fc.SetDataStore(this);
                    }
                    _rowIndexer.Set(cell);
                }
            }
        }
    }
}
