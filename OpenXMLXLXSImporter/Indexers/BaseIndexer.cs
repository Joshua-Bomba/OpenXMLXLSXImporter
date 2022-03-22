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
        public abstract bool TryGetCell(uint rowIndex, string cellIndex, out ICellIndex ci);
        public virtual async Task ProcessInstruction(ISpreadSheetInstruction instruction)
        {
            using (await _lock.IndexerLock.LockAsync())
            {
                instruction.EnqueCell(this);
            }
        }

        public void Spread(ICellIndex cell)
        {
            InternalAdd(cell);
        }

        void IIndexer.Add(ICellProcessingTask index)
        {
            if(index is ICellIndex c)
            {
                InternalAdd(c);
                _lock.Spread(this, c);
            }

            _lock.EnqueCell(index);
        }
    }
}
