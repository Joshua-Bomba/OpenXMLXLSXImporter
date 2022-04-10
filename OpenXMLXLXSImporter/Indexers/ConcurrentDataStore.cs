using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Processing;
using OpenXMLXLSXImporter.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public class ConcurrentDataStore : IFutureUpdate
    {
        protected DirectDataStore _rowIndexer;
        private AsyncLock _accessorLock = new AsyncLock();
        public ConcurrentDataStore(IQueueAccess queue)
        {
            _rowIndexer = new DirectDataStore(queue,this);
        }
        public async Task<LastColumn> GetLastColumn(uint rowIndex)
        {
            if(!_rowIndexer.ContainsKey(rowIndex)||_rowIndexer[rowIndex].LastColumn == null)
            {
                using (await _accessorLock.LockAsync())
                {
                   return _rowIndexer.GetLastColumn(rowIndex);
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
                    _rowIndexer.GetLastRow();
                }
            }
            return _rowIndexer.LastRow;
        }


        public async Task<ICellIndex> GetCell(uint rowIndex, string cellIndex)
        {
            ICellIndex result = _rowIndexer.Get(rowIndex, cellIndex);
            if (result == null)
            {
                using (await this._accessorLock.LockAsync())
                {
                    return this._rowIndexer.GetCell(rowIndex, cellIndex);
                }
            }
            return result;
        }

        public virtual async Task ProcessInstruction(ISpreadSheetInstruction instruction)
        {
            using (await _accessorLock.LockAsync())
            {
                LimitedAccessDataStore limitedLife = new LimitedAccessDataStore(_rowIndexer);
                try
                {
                    instruction.EnqueCell(limitedLife);
                }
                finally
                {
                    limitedLife.Delete();
                }

            }
        }

        public virtual async Task ProcessInstructions(IEnumerable<ISpreadSheetInstruction> instructions)
        {
            using(await _accessorLock.LockAsync())
            {
                LimitedAccessDataStore limitedLife = new LimitedAccessDataStore(_rowIndexer);
                try
                {
                    foreach (ISpreadSheetInstruction instruction in instructions)
                    {
                        instruction.EnqueCell(limitedLife);
                    }
                }
                finally
                {
                    limitedLife.Delete();
                }
            }
        }

        public async Task AddDeferredCells(IEnumerable<DeferredCell> cells)
        {
            using (await _accessorLock.LockAsync())
            {
                this._rowIndexer.AddDeferredCells(cells);
            }
        }

        //we don't want to wait for this, we will update it when we get around to it
        void IFutureUpdate.Update(ICellIndex cell) => Task.Run(async () =>
        {
            using (await _accessorLock.LockAsync())
            {
                this._rowIndexer.Set(cell);
            }
        });
    }
}
