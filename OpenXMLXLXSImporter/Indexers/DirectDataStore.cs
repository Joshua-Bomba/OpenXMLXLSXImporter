using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Processing;
using OpenXMLXLSXImporter.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public class DirectDataStore : Dictionary<uint,ColumnIndexer>, IDataStore
    {
        private IQueueAccess _queueAccess;
        private IFutureUpdate _futureUpdate;
        public DirectDataStore(IQueueAccess queueAccess, IFutureUpdate updater)
        {
            _queueAccess = queueAccess;
            _futureUpdate = updater;
        }

        public LastRow LastRow { get; set; }

        public  ICellIndex Get(uint rowIndex, string cellIndex)
        {
            if (base.ContainsKey(rowIndex) && base[rowIndex].ContainsKey(cellIndex))
            {
                return base[rowIndex][cellIndex];
            }
            return null;
        }

        public void Set(ICellIndex cellData)
        {
            if (!base.ContainsKey(cellData.CellRowIndex))
            {
                base.Add(cellData.CellRowIndex, new ColumnIndexer());
            }
            base[cellData.CellRowIndex][cellData.CellColumnIndex] = cellData;
        }

        public void Add(ICellIndex cellData)
        {
            if (!base.ContainsKey(cellData.CellRowIndex))
            {
                base.Add(cellData.CellRowIndex, new ColumnIndexer());
            }
            base[cellData.CellRowIndex].Add(cellData.CellColumnIndex, cellData);
        }


        public async Task ProcessInstruction(ISpreadSheetInstruction instruction)
        {
            await instruction.EnqueCell(this);
        }

        public async Task<ICellIndex> GetCell(uint rowIndex, string cellIndex, Func<ICellProcessingTask> newCell = null)
        {
            ICellIndex r = this.Get(rowIndex, cellIndex);
            if (r == null&&newCell != null)
            {
                await _queueAccess.LockQueue(x =>
                {
                    r = this.Get(rowIndex, cellIndex);
                    if (r == null && newCell != null)
                    {
                        ICellProcessingTask t = newCell();
                        if (t != null)
                        {
                            if (t is IFutureCell fc)
                            {
                                fc.Updateder(_futureUpdate);
                            }
                            x.Enque(t);
                            if (t is ICellIndex index)
                            {
                                r = index;
                                this.Set(index);
                            }
                        }
                    }
                });
            }
            return r;
        }

        public async Task<LastColumn> GetLastColumn(uint rowIndex)
        {
            if (!this.ContainsKey(rowIndex))
            {
                this[rowIndex] = new ColumnIndexer();
            }
            if (this[rowIndex].LastColumn == null)
            {
                this[rowIndex].LastColumn = new LastColumn(rowIndex);
                await this._queueAccess.QueueNonIndexedCell(this[rowIndex].LastColumn);
            }
            return this[rowIndex].LastColumn;
        }

        public async Task<LastRow> GetLastRow()
        {
            if (this.LastRow == null)
            {
                this.LastRow = new LastRow();
                await this._queueAccess.QueueNonIndexedCell(this.LastRow);
            }
            return this.LastRow;
        }

        public async Task SetCell(ICellIndex index)
        {
            if (index is IFutureCell fc)
            {
                fc.Updateder(_futureUpdate);
            }
            this.Set(index);
        }

        public async Task SetCells(IEnumerable<ICellIndex> cells)
        {
            foreach (ICellIndex cell in cells)
            {
                if (cell is IFutureCell fc)
                {
                    fc.Updateder(_futureUpdate);
                }
                this.Set(cell);
            }
        }
    }
}
