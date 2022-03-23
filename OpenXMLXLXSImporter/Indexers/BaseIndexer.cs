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
        protected ISpreadSheetIndexersLock _lock;
        public BaseIndexer(ISpreadSheetIndexersLock indexerLock)
        {
            _lock = indexerLock;
            _lock.AddIndexer(this);
        }



        protected abstract void InternalAdd(ICellIndex cell);
        protected abstract void InternalSet(ICellIndex cell);
        public abstract bool TryGetCell(uint rowIndex, string cellIndex, out ICellIndex ci);
        public virtual async Task ProcessInstruction(ISpreadSheetInstruction instruction)
        {
            using (await _lock.IndexerLock.LockAsync())
            {
                await instruction.EnqueCell(this);
            }
        }

        public void Spread(ICellIndex cell)
        {
            InternalSet(cell);
        }

        async Task IIndexer.Add(ICellProcessingTask index)
        {
            if(index is IFutureCell fc)
            {
                fc.SetIndexer(this);
            }

            if(index is ICellIndex c)
            {
                InternalAdd(c);
                _lock.Spread(this, c);
            }

           await _lock.EnqueCell(index);
        }

        async Task IIndexer.SetCell(ICellData d)
        {
            using(await _lock.IndexerLock.LockAsync())
            {
                InternalSet(d);
                _lock.Spread(this, d);
            }
        }
    }
}
