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
        public async Task<ICellIndex> GetLastColumn(uint rowIndex)
        {
            throw new NotImplementedException();
        }

        public async Task<ICellIndex> GetLastRow(uint columnIndex)
        {
            throw new NotImplementedException();
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
                                fc.SetIndexer(this);
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

        //async Task IIndexer.Add(ICellProcessingTask index)
        //{
        //    if(index is IFutureCell fc)
        //    {
        //        fc.SetIndexer(this);
        //    }

        //    if(index is ICellIndex c)
        //    {
        //        InternalAdd(c);
        //        _lock.Spread(this, c);
        //    }

        //   await _lock.EnqueCell(index);
        //}

        public async Task SetCell(ICellIndex index)
        {
            using(await _accessorLock.LockAsync())
            {
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

            }
        }
    }
}
