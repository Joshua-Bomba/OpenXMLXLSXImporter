﻿using DocumentFormat.OpenXml;
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

    public class SpreadSheetInstructionManager : IQueueAccess, ISpreadSheetInstructionManager
    {       
        private ConcurrentDataStore _dataStore;

        private ChunkableBlockingCollection<ICellProcessingTask> _loadQueueManager;

        public IQueueAccess Queue => this;

        public SpreadSheetInstructionManager(SpreadSheetDequeManager dequeManager)
        {
            _loadQueueManager = new ChunkableBlockingCollection<ICellProcessingTask>(dequeManager);
            _dataStore = new ConcurrentDataStore(this);
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
            }
            await _dataStore.SetCells(cells);         
        }

        public async Task QueueNonIndexedCell(ICellProcessingTask t)
        {
            using(await _loadQueueManager.Mutex.LockAsync())
            {
                _loadQueueManager.Enque(t);
            }
        }

        public async Task LockQueue(Action<IChunkableBlockingCollection<ICellProcessingTask>> lockedQueue)
        {
            using(await _loadQueueManager.Mutex.LockAsync())
            {
                lockedQueue(_loadQueueManager);
            }
        }
    }
}
