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
    public class SpreadSheetDequeManager
    {
        private IDeferredUpdater _deferredUpdater;

        private AsyncProducerConsumerQueue<ICellProcessingTask> _queue;

        private uint? desiredRowIndex;
        private string desiredColumnIndex;
        private Dictionary<string,Cell> deferedCells = null;
        private IXlsxSheetFilePromise _filePromise;
        private IXlsxSheetFile sheetAccess;


        public SpreadSheetDequeManager(IDeferredUpdater deferredUpdater)
        {
            _deferredUpdater = deferredUpdater;
            _filePromise = null;
            _queue = new AsyncProducerConsumerQueue<ICellProcessingTask>();
        }

        public void Finish() => _queue.CompleteAdding();

        public async Task QueueAsync(ICellProcessingTask queue)
        {
            await _queue.EnqueueAsync(queue);
        }

        public void Terminate(Exception e)
        {
            _queue.CompleteAdding();
            ICellProcessingTask[] tasks = _queue.GetConsumingEnumerable().ToArray();
            foreach(ICellProcessingTask t in tasks)
            {
                t.Failure(e);
            }
        }

        public async Task ProcessRequests(IXlsxSheetFilePromise file)
        {
            _filePromise = file;
            try
            {
                Cell cell = null;
                bool rowsLoadedIn = false;
                IEnumerator<Cell> cellEnumerator;
                sheetAccess = await _filePromise.GetLoadedFile();
                ICellIndex index;
                while (true)
                {
                    await ProcessDeferredCells();
                    ICellProcessingTask dequed = await _queue.DequeueAsync();
                    if(!dequed.Processed)
                    {
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
                            if (desiredRowIndex == null)
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
                                if (currentIndex != desiredColumnIndex)
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
            }
            catch (InvalidOperationException ex)
            {
                this.Terminate(ex);
            }
        }

        public async Task ProcessDeferredCells()
        {
            //We will add this deferredcell type to the IIndexers since we don't need them at the time
            Task<Dictionary<DeferredCell,ICellProcessingTask>> setDeferedCells = null;

            if (deferedCells != null && deferedCells.Any())
            {
                setDeferedCells = _deferredUpdater.AddDeferredCells(deferedCells.Select(x => new DeferredCell(desiredRowIndex.Value, x.Key, x.Value)).ToArray());
            }

            if(setDeferedCells != null)
            {
                Dictionary<DeferredCell, ICellProcessingTask> cells = await setDeferedCells;
                foreach(KeyValuePair<DeferredCell,ICellProcessingTask> kv in cells)
                {
                    if (!kv.Value.Processed)
                    {
                        kv.Value.Resolve(sheetAccess, kv.Key.Cell, kv.Key);
                    }
                }
                deferedCells = null;
            }
        }
    }
}
