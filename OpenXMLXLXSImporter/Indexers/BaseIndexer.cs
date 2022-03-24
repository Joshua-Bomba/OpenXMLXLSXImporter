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
    public abstract class BaseIndexer : IIndexer
    {
        protected SpreadSheetInstructionManager _instructionManager;
        public BaseIndexer(SpreadSheetInstructionManager instructionManager)
        {
            _instructionManager = instructionManager;
            _instructionManager.AddIndexer(this);
        }

        protected abstract ICellIndex InternalGet(uint rowIndex, string cellIndex);

        protected abstract void InternalAdd(ICellIndex cell);
        protected abstract void InternalSet(ICellIndex cell);
        public async Task<ICellIndex> GetCell(uint rowIndex, string cellIndex, Func<ICellProcessingTask> newCell = null)
        {
            ICellIndex result = InternalGet(rowIndex, cellIndex);
            if (result == null&&newCell != null)
            {
                using (await this._instructionManager.Queue.Mutex.LockAsync())
                {
                    ICellIndex r = InternalGet(rowIndex, cellIndex);
                    if (r == null)
                    {
                        ICellProcessingTask t = newCell();
                        if (t != null)
                        {
                            this._instructionManager.Queue.Enque(t);
                        }
                    }
                    return r;
                }
            }
            return result;
        }
        public virtual async Task ProcessInstruction(ISpreadSheetInstruction instruction)
        {
            using (await _instructionManager.IndexerLock.LockAsync())
            {
                await instruction.EnqueCell(this);
            }
        }

        public void Spread(ICellIndex cell)
        {
            InternalSet(cell);
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

        async Task IIndexer.SetCell(ICellData d)
        {
            using(await _instructionManager.IndexerLock.LockAsync())
            {
                InternalSet(d);
                _instructionManager.Spread(this, d);
            }
        }

        public async Task QueueNonIndexedCell(ICellProcessingTask t)
        {
            using (await _instructionManager.Queue.Mutex.LockAsync())
            {
                _instructionManager.Queue.Enque(t);
            }
        }
    }
}
