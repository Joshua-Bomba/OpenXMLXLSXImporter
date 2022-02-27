using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
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
        private ExcelImporter _importer;

        private ChunkableBlockingCollection<ICellProcessingTask> _queue;

        private uint desiredRowIndex;
        private Dictionary<string,Cell> deferedCells = null;
        private IXlsxSheetFilePromise _filePromise;
        private IXlsxSheetFile sheetAccess;

        private SortedSet<ICellProcessingTask> ss;
        private Dictionary<Cell, ICellProcessingTask> fufil;


        private Task _processRequestTask;
        public SpreadSheetDequeManager(ExcelImporter importer)
        {
            _importer = importer;
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

                while (true)
                {
                    ICellProcessingTask t = _queue.Take();
                    desiredRowIndex = t.CellRowIndex;
                    if (sheetAccess.TryGetRow(desiredRowIndex, out cellEnumerator))
                    {
                        string columnIndex = t.CellColumnIndex;
                        bool cellsLoadedIn = false;
                        string currentIndex;
                        do
                        {
                            if (cellEnumerator.MoveNext())
                            {
                                cell = cellEnumerator.Current;
                                currentIndex = XlsxSheetFile.GetColumnIndexByColumnReference(cell.CellReference);
                                if (currentIndex != columnIndex)
                                {
                                    if (deferedCells == null)
                                        deferedCells = new Dictionary<string, Cell>();
                                    deferedCells.Add(currentIndex,cell);//this does not actually do anything since everything we need is loaded in at this point in our current version of the code. that's the first thing we should update
                                    //also we should make the enque and fetch steps occur at the same time
                                }
                            }
                            else
                            {
                                currentIndex = null;
                                cellsLoadedIn = true;
                            }
                        } while (!cellsLoadedIn && currentIndex != columnIndex);

                        if (currentIndex == columnIndex)
                        {
                            sheetAccess.ProcessedCell(cell, t);
                        }
                        else
                        {
                            //This Cell Does not exist
                            t.Resolve(null);
                        }

                    }
                    else
                    {
                        //if we reached the end of the file and the row does not exist then what were trying to get does not exist
                        t.Resolve(null);
                    }
                }
            }
            catch (InvalidOperationException ex)
            {
                //queue is finished
            }
        }

        public bool ShouldPullAndChunk => deferedCells?.Any()??false;

        public bool KeepQueueLockedForDump() => true;

        private class DequeOrder : IComparer<ICellProcessingTask>
        {
            public int Compare(ICellProcessingTask x, ICellProcessingTask y)
            {
                return x.CellColumnIndex.CompareTo(y.CellColumnIndex);
            }
        }
        public void PreProcessing()
        {
            ss = new SortedSet<ICellProcessingTask>(new DequeOrder());
            fufil = new Dictionary<Cell, ICellProcessingTask>();
        }

        public void QueueDumpped(ref List<ICellProcessingTask> items)
        {

            foreach(ICellProcessingTask item in items)
            {
                if (item.CellRowIndex == desiredRowIndex&&deferedCells.ContainsKey(item.CellColumnIndex))
                {
                    fufil.Add(deferedCells[item.CellColumnIndex], item);
                    deferedCells.Remove(item.CellColumnIndex);
                }
                else
                {
                    ss.Add(item);
                }
            }
            items = ss.ToList();

            //We will add this deferredcell type to the IIndexers since we don't need them at the time
            foreach (KeyValuePair<string,Cell> cell in deferedCells)
            {
                DeferredCell dc = new DeferredCell(desiredRowIndex,cell.Key,cell.Value);
                _importer.Spread(dc);
            }
        }

        public void PostProcessing()
        {
            deferedCells = null;
            ss = null;
            //fufill any cells that were enqued during the processing of the last add
            foreach(KeyValuePair<Cell,ICellProcessingTask> kv in fufil)
            {
                sheetAccess.ProcessedCell(kv.Key, kv.Value);
            }
            fufil = null;
        }

    }
}
