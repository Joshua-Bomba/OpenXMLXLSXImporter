using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.FileAccess;
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
    public class DirectDataStore : Dictionary<uint,ColumnIndexer>
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

        public ICellIndex GetCell(uint rowIndex, string cellIndex)
        {
            ICellIndex r = this.Get(rowIndex, cellIndex);
            if (r == null)
            {
                FutureCell fc = new FutureCell(rowIndex, cellIndex,_futureUpdate, _queueAccess);
                if (fc != null)
                {
                    r = fc;
                    this.Set(fc);
                }
            }
            return r;
        }

        public LastColumn GetLastColumn(uint rowIndex)
        {
            if (!this.ContainsKey(rowIndex))
            {
                this[rowIndex] = new ColumnIndexer();
            }
            if (this[rowIndex].LastColumn == null)
            {
                this[rowIndex].LastColumn = new LastColumn(rowIndex, this._queueAccess);
            }
            return this[rowIndex].LastColumn;
        }

        public LastRow GetLastRow()
        {
            if (this.LastRow == null)
            {
                this.LastRow = new LastRow(this._queueAccess);
            }
            return this.LastRow;
        }

        public Dictionary<DeferredCell,ICellProcessingTask> AddDeferredCells(IEnumerable<DeferredCell> dc)
        {
            Dictionary<DeferredCell, ICellProcessingTask> existing = new Dictionary<DeferredCell, ICellProcessingTask>();
            foreach (DeferredCell c in dc)
            {
                ICellIndex i = this.Get(c.CellRowIndex, c.CellColumnIndex);
                if (i != null)
                {
                    if (i is ICellProcessingTask task)
                    {
                        existing.Add(c, task);
                    }
                    //if i is not ICellProcessingTask it means it's already resolved in which case we will just dump the deferredcell
                    //I don't think this is possible but with the rube goldberg machine this is anything is possible
                    //maybe is was manually set
                }
                else
                {
                    c.Updater = _futureUpdate;
                    c.QueueAccess = _queueAccess;
                    this.Set(c);
                }
            }
            return existing;
        }
    }
}
