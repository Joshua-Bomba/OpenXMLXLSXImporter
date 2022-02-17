using Nito.AsyncEx;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Utils;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Processing
{
    public class SpreadSheetDequeManager : IChunckBlock<ICellProcessingTask>
    {
        private Queue<DeferredCell> deferedCells = null;
        public SpreadSheetDequeManager()
        {
            
        }

        public void AddDeferredCell(DeferredCell deferredCell)
        {
            if (deferedCells == null)
                deferedCells = new Queue<DeferredCell>();
            deferedCells.Enqueue(deferredCell);
        }

        public bool ShouldPullAndChunk => false;

        public bool KeepQueueLockedForDump()
        {
            return false;
        }

        public void QueueDumpped(ref List<ICellProcessingTask> items)
        {
            
        }
    }
}
