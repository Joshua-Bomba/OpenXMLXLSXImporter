using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.FileAccess;
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
        private ISpreadSheetInstructionManager _instructionManager;

        private ChunkableBlockingCollection<ICellProcessingTask> _queue;

        private uint? desiredRowIndex;
        private string desiredColumnIndex;
        private Dictionary<string,Cell> deferedCells = null;
        private IXlsxSheetFilePromise _filePromise;
        private IXlsxSheetFile sheetAccess;

        private Queue<ICellProcessingTask> ss;
        private Dictionary<Cell, ICellProcessingTask> fufil;


        private Task _processRequestTask;
        public SpreadSheetDequeManager(ISpreadSheetInstructionManager instructionManager)
        {
            _instructionManager = instructionManager;
            _filePromise = null;
            _processRequestTask = null;
        }

        public void Finish() => _queue.Finish();

        public void Init(ChunkableBlockingCollection<ICellProcessingTask> collection)
        {
            _queue = collection;
        }

        public void StartRequestProcessor(IXlsxSheetFilePromise file)
        {
            if (_processRequestTask == null)
            {
                _filePromise = file;
                _processRequestTask = Task.Run(ProcessRequests);
            }

        }

        protected async Task ProcessRequests()
        {
            try
            {
                Cell cell = null;
                bool rowsLoadedIn = false;
                IEnumerator<Cell> cellEnumerator;
                sheetAccess = await _filePromise.GetLoadedFile();
                ICellIndex index;
                while (true)
                {
                    ICellProcessingTask dequed = await _queue.Take();
                    index = null;
                    cell = null;
                    desiredRowIndex = null;
                    try
                    {
                        if (dequed is ICellIndex t)
                        {
                            index = t;
                            desiredRowIndex = t.CellRowIndex;
                            desiredColumnIndex = t.CellColumnIndex;
                        }
                        else if (dequed is LastRow m)
                        {
                            index = new FutureIndex { CellRowIndex = sheetAccess.GetAllRows() };
                            continue;
                        }
                        else if (dequed is LastColumn mc)
                        {
                            index = null;
                            desiredRowIndex = mc._row;
                            desiredColumnIndex = null;
                        }
                        else
                        {
                            
                            continue;
                        }
                        if(desiredRowIndex == null)
                        {
                            continue;
                        }
                        if (sheetAccess.TryGetRow(desiredRowIndex.Value, out cellEnumerator))
                        {
                            bool cellsLoadedIn = false;
                            string currentIndex = null;
                            do
                            {
                                if (cellEnumerator.MoveNext())
                                {
                                    cell = cellEnumerator.Current;
                                    currentIndex = XlsxSheetFile.GetColumnIndexByColumnReference(cell.CellReference);
                                    if (currentIndex != desiredColumnIndex)
                                    {
                                        if (deferedCells == null)
                                            deferedCells = new Dictionary<string, Cell>();
                                        deferedCells.Add(currentIndex, cell);//this does not actually do anything since everything we need is loaded in at this point in our current version of the code. that's the first thing we should update
                                                                             //also we should make the enque and fetch steps occur at the same time
                                    }
                                }
                                else
                                {
                                    if (index == null)
                                    {
                                        index = new FutureIndex { CellColumnIndex = currentIndex, CellRowIndex = desiredRowIndex.Value };
                                    }
                                    currentIndex = null;
                                    cellsLoadedIn = true;
                                }
                            } while (!cellsLoadedIn && currentIndex != desiredColumnIndex);
                            if(currentIndex != desiredColumnIndex)
                            {
                                cell = null;
                            }
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                    finally
                    {
                        dequed.Resolve(sheetAccess, cell, index);
                    }
                }
            }
            catch (InvalidOperationException ex)
            {
                //queue is finished
            }
        }

        public bool ShouldPullAndChunk => deferedCells?.Any() ?? false;

        public async Task PreLockProcessing()
        {
            ss = new Queue<ICellProcessingTask>();
            fufil = new Dictionary<Cell, ICellProcessingTask>();
        }

        public async Task PreQueueProcessing()
        {
            
        }

        public void ProcessQueue(ref Queue<ICellProcessingTask> items)
        {
            ICellProcessingTask[] tasks = items.ToArray();
            List<ICellIndex> indexes = new List<ICellIndex>(tasks.Length);
            int reuse = 0;
            ICellProcessingTask task;
            for (int i = 0; i < tasks.Length; i++)
            {
                task = tasks[i];
                tasks[i] = null;
                if (task is ICellIndex item)
                {
                    if (deferedCells != null&&desiredRowIndex.HasValue&&item.CellRowIndex == desiredRowIndex.Value&&deferedCells.ContainsKey(item.CellColumnIndex))
                    {
                        fufil.Add(deferedCells[item.CellColumnIndex], task);
                        deferedCells.Remove(item.CellColumnIndex);
                    }
                    else
                    {
                        indexes.Add(item);
                    }
                }
                else
                {
                    tasks[reuse] = task;
                    reuse++;
                }
            }

            IOrderedEnumerable<IGrouping<uint, ICellIndex>> g = indexes.GroupBy(x => x.CellRowIndex).OrderBy(x => x.Key);
            foreach (IGrouping<uint, ICellIndex> row in g)
            {
                IOrderedEnumerable<ICellIndex> currentRow = row.OrderBy(x => ExcelColumnHelper.GetColumnStringAsIndex(x.CellColumnIndex));
                foreach (ICellIndex item in currentRow)
                {
                    ss.Enqueue(item as ICellProcessingTask);
                }
            }

            for (int i = 0; i < reuse; i++)
            {
                ss.Enqueue(tasks[i]);
            }

            items = ss;
           
        }

        public async Task PostQueueProcessing()
        {
            //We will add this deferredcell type to the IIndexers since we don't need them at the time
            if(deferedCells != null&&deferedCells.Any())
            {
                await _instructionManager.AddDeferredCells(deferedCells.Select(x => new DeferredCell(desiredRowIndex.Value, x.Key, x.Value)));
            }
        }

        public async Task PostLockProcessing()
        {
            deferedCells = null;
            ss = null;
            //fufill any cells that were enqued during the processing of the last add
            foreach (KeyValuePair<Cell, ICellProcessingTask> kv in fufil)
            {
                kv.Value.Resolve(sheetAccess, kv.Key, kv.Value as ICellIndex);
            }
            fufil = null;
        }
    }
}
