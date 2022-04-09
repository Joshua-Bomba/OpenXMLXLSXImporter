using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.FileAccess;
using OpenXMLXLSXImporter.Indexers;
using OpenXMLXLSXImporter.Utils;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Processing
{
    public interface IQueueAccess
    {
        Task QueueNonIndexedCell(ICellProcessingTask t);

        Task LockQueue(Action<IChunkableBlockingCollection<ICellProcessingTask>> lockedQueue);
    }

    public interface ISpreadSheetInstructionManager
    {
        Task AddDeferredCells(IEnumerable<DeferredCell> deferredCells);
        Task ProcessInstruction(ISpreadSheetInstruction spreadSheetInstruction);
        IQueueAccess Queue { get; }//Did it like this so that way I can possbily replace IQueueAccess with another implementation
    }

    public class SpreadSheetInstructionManager : IQueueAccess, ISpreadSheetInstructionManager, IDisposable
    {
        private ConcurrentDataStore _dataStore;

        private SpreadSheetDequeManager dequeManager;
        private ChunkableBlockingCollection<ICellProcessingTask> _loadQueueManager;

        private AsyncManualResetEvent _queueInit;
        private Task _instructionProcessor;
        public IQueueAccess Queue => this;

        public SpreadSheetInstructionManager(Task<IXlsxSheetFilePromise> sheetFilePromise)
        {
            _dataStore = new ConcurrentDataStore(this);
            _queueInit = new AsyncManualResetEvent(false);
            ProcessSheet(sheetFilePromise);
        }

        private void ProcessSheet(Task<IXlsxSheetFilePromise> sheetFilePromise)
        {
            _instructionProcessor = Task.Run(async () =>
            {
                dequeManager = new SpreadSheetDequeManager(this);
                _loadQueueManager = new ChunkableBlockingCollection<ICellProcessingTask>(dequeManager);
                _queueInit.Set();
                IXlsxSheetFilePromise g = await sheetFilePromise;
                await dequeManager.ProcessRequests(g);
            });
        }
        public void Dispose()
        {
            _queueInit.Wait();
            dequeManager.Finish();
            _instructionProcessor.Wait();
        }

        public async Task ProcessInstructionBundle(IEnumerable<ISpreadSheetInstruction> bundle)
        {
            ISpreadSheetInstruction[] instructions = bundle.ToArray();
            foreach(ISpreadSheetInstruction instruction in instructions)
            {
                instruction.AttachSpreadSheetInstructionManager(this);
            }
            await _dataStore.ProcessInstructions(instructions);
        }

        public async Task ProcessInstruction(ISpreadSheetInstruction spreadSheetInstruction)
        {
            spreadSheetInstruction.AttachSpreadSheetInstructionManager(this);
            await _dataStore.ProcessInstruction(spreadSheetInstruction);
        }

        public  async Task AddDeferredCells(IEnumerable<DeferredCell> deferredCells)
        {
            DeferredCell[] cells = deferredCells.ToArray();
            for(int i =0;i < cells.Length;i++)
            {
                cells[i].InstructionManager = this;
                cells[i].Updater = _dataStore;
            }
            await _dataStore.SetCells(cells);         
        }

        public async Task QueueNonIndexedCell(ICellProcessingTask t)
        {
            await _queueInit.WaitAsync();
            using (await _loadQueueManager.Mutex.LockAsync())
            {
                _loadQueueManager.Enque(t);
            }
        }

        public async Task LockQueue(Action<IChunkableBlockingCollection<ICellProcessingTask>> lockedQueue)
        {
            await _queueInit.WaitAsync();
            using (await _loadQueueManager.Mutex.LockAsync())
            {
                lockedQueue(_loadQueueManager);
            }
        }
    }
}
