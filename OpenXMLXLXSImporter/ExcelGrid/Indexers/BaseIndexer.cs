using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
{
    public abstract class BaseIndexer : IIndexer
    {
        protected ISpreadSheetIndexersLock _lock;
        public BaseIndexer(ISpreadSheetIndexersLock indexerLock)
        {
            _lock = indexerLock;
            _lock.AddIndexer(this);
        }

        protected abstract void InternalAdd(ICellIndex cell);

        protected abstract bool InternalContains(uint rowIndex, string cellIndex);
        protected abstract ICellIndex InternalGet(uint rowIndex, string cellIndex);

        public virtual async Task Add(ICellIndex cellData)
        {
            using (await _lock.IndexerLock.LockAsync())
            {
                InternalAdd(cellData);
            }
        }

        public async Task<ICellData> GetCell(uint rowIndex, string cellIndex)
        {
            ICellIndex i;
            using (await _lock.IndexerLock.LockAsync())
            {
              i = InternalGet(rowIndex, cellIndex);
            }
            if (i is ICellData cd)
            {
                return cd;
            }
            if (i is IFutureCell fc)
            {
                return await fc.GetCell();
            }
            throw new InvalidOperationException();
        }

        public async Task<bool> HasCell(uint rowIndex, string cellIndex)
        {
            using (await _lock.IndexerLock.LockAsync())
            {
                return InternalContains(rowIndex, cellIndex);
            }
        }

        public void Spread(ICellIndex cell)
        {
            InternalAdd(cell);
        }
    }
}
