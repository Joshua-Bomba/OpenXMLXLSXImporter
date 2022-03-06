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

        public abstract bool HasCell(uint rowIndex, string cellIndex);
        public abstract ICellIndex GetCell(uint rowIndex, string cellIndex);

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

        public void Add(ICellProcessingTask index)
        {
            InternalAdd(index);
            _lock.Spread(this, index);
            _lock.EnqueCell(index);
        }
    }
}
