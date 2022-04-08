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
    /// <summary>
    /// this will handle accessing the grid and iterateing over the spreadsheetgrid
    /// interate through all the cells for an entire column
    /// interate through all the cells for an entire row
    /// interate through all the cells
    /// </summary>
    public class SpreadSheetInstructionManager
    {       
        private DataStore _dataStore;


        private ChunkableBlockingCollection<ICellProcessingTask> _loadQueueManager;

        public AsyncLock IndexerLock => _accessorLock;

        public IChunkableBlockingCollection<ICellProcessingTask> Queue => _loadQueueManager;

        public SpreadSheetInstructionManager(SpreadSheetDequeManager dequeManager)
        {
            _loadQueueManager = new ChunkableBlockingCollection<ICellProcessingTask>(dequeManager);
            _dataStore = new DataStore(this);
        }

        public async Task ProcessInstruction(ISpreadSheetInstruction spreadSheetInstruction)
        {
            await _dataStore.ProcessInstruction(spreadSheetInstruction);
        }

        public  async Task AddDeferredCells(IEnumerable<DeferredCell> deferredCells)
        {
            DeferredCell[] cells = deferredCells.ToArray();
            for(int i =0;i < cells.Length;i++)
            {
                cells[i].InstructionManager = this;
            }
            using(await IndexerLock.LockAsync())
            {
                foreach (DeferredCell deferredCell in cells)
                {
                    deferredCell.SetIndexer(_dataStore);
                    _dataStore.SetCell(deferredCell);
                }
            }            
        }



        //public async Task Add(ICellData cellData)
        //{
        //    using(await _lockRow.LockAsync())
        //    {
        //        if (!_rows.ContainsKey(cellData.CellRowIndex))
        //        {
        //            _rows[cellData.CellRowIndex] = new RowIndexer();
        //        }

        //        _rows[cellData.CellRowIndex].Add(cellData);
        //        if(_listeners != null)
        //        {
        //            //intresting so this will call each method and won't await till it's finished calling all of them
        //            //pretty handy neat pattern
        //            await Task.WhenAll(_listeners.Select(x => x.NotifyAsync(cellData)));
        //        }
        //    }
        //    using (await _lockColumn.LockAsync())
        //    {
        //        if (!_columns.ContainsKey(cellData.CellColumnIndex))
        //        {
        //            _columns[cellData.CellColumnIndex] = new ColumnIndexer(this);
        //        }

        //        _columns[cellData.CellColumnIndex].Add(cellData);
        //    }
        //}
    }
}
