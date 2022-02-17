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
        private ChunkableBlockingCollection<ICellProcessingTask> _queue;

        private Queue<DeferredCell> deferedCells = null;
        private IXlsxSheetFilePromise _filePromise;

        private Task _processRequestTask;
        public SpreadSheetDequeManager()
        {
            _filePromise = null;
            _processRequestTask = null;
        }

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
                Cell cell;
                uint desiredRowIndex;
                bool rowsLoadedIn = false;
                IEnumerator<Cell> cellEnumerator;
                IXlsxSheetFile sheetAccess = await _filePromise.GetLoadedFile();

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
                                currentIndex = XlsxDocumentFile.GetColumnIndexByColumnReference(cell.CellReference);
                                if (currentIndex != columnIndex)
                                {
                                    this.AddDeferredCell(new DeferredCell(desiredRowIndex, currentIndex, cell));
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
                            throw new NotImplementedException();//need to handle this step now which is loading in the actual data
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
