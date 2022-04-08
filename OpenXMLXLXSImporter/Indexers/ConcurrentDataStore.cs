﻿using Nito.AsyncEx;
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




    public class ConcurrentDataStore : IDataStore
    {
        protected DirectDataStore _rowIndexer;
        private AsyncLock _accessorLock = new AsyncLock();
        public ConcurrentDataStore(IQueueAccess queue)
        {
            _rowIndexer = new DirectDataStore(queue);
        }
        public async Task<LastColumn> GetLastColumn(uint rowIndex)
        {
            if(!_rowIndexer.ContainsKey(rowIndex)||_rowIndexer[rowIndex].LastColumn == null)
            {
                using (await _accessorLock.LockAsync())
                {
                   return await _rowIndexer.GetLastColumn(rowIndex);
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
                    await _rowIndexer.GetLastRow();
                }
            }
            return _rowIndexer.LastRow;
        }


        public async Task<ICellIndex> GetCell(uint rowIndex, string cellIndex, Func<ICellProcessingTask> newCell = null)
        {
            ICellIndex result = _rowIndexer.Get(rowIndex, cellIndex);
            if (result == null&&newCell != null)
            {
                using (await this._accessorLock.LockAsync())
                {
                    return await this._rowIndexer.GetCell(rowIndex, cellIndex, newCell);
                }
            }
            return result;
        }


        public virtual async Task ProcessInstruction(ISpreadSheetInstruction instruction)
        {
            using (await _accessorLock.LockAsync())
            {
                await this._rowIndexer.ProcessInstruction(instruction);
            }
        }

        public async Task SetCell(ICellIndex index)
        {
            using(await _accessorLock.LockAsync())
            {
                await this._rowIndexer.SetCell(index);
            }
        }

        public async Task SetCells(IEnumerable<ICellIndex> cells)
        {
            using (await _accessorLock.LockAsync())
            {
                await this._rowIndexer.SetCells(cells);
            }
        }
    }
}