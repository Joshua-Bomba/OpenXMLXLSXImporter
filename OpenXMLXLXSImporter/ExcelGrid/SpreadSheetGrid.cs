using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.Enumerators;
using OpenXMLXLXSImporter.ExcelGrid.Indexers;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid
{
    /// <summary>
    /// this will handle accessing the grid and iterateing over the spreadsheetgrid
    /// interate through all the cells for an entire column
    /// interate through all the cells for an entire row
    /// interate through all the cells
    /// </summary>
    public class SpreadSheetGrid : IAsyncEnumerable<ICellData>, ISpreadSheetGridCollectionAccessor
    {
        ISheetProperties _sheet;
        private Dictionary<uint, RowIndexer> _rows;
        private Dictionary<string, ColumnIndexer> _columns;
        private bool _allLoaded;
        private readonly AsyncLock _lockRow;
        private readonly AsyncLock _lockColumn;
        private List<IItemEnquedEvent> _listeners;
        public SpreadSheetGrid(ISheetProperties sheet)
        {
            _sheet = sheet;
            sheet.SetCellGrid(this);
            _allLoaded = false;
            _rows = new Dictionary<uint, RowIndexer>();
            _columns = new Dictionary<string, ColumnIndexer>();
            _lockRow = new AsyncLock();
            _lockColumn = new AsyncLock();
            _listeners = null;
        }
        /// <summary>
        /// this will return a paticular cell if it's avaliable
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public bool TryFetchCell(uint row, string cell, out ICellData cellData)
        {
            cellData = null;
            if (_rows.ContainsKey(row))
            {
                return _rows[row].TryAndGetCell(cell, out cellData);
            }
            return false;
        }

        /// <summary>
        /// this will return a paticular row and cell if it's not avaliable it will wait till it's loaded or the timeout is reached
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="cellIndex"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public async Task<ICellData> FetchCell(uint rowIndex, string cellIndex, int timeout = -1)
        {
            ICellData cell = null;
            await Task.Run(async() =>
            {
                ManualResetEvent e = null;
                using(await _lockRow.LockAsync())//honestly this multi threading crap is probably unessary unless you havea massive excel document but whatever this is how I wrote it
                {
                    if (!_rows.ContainsKey(rowIndex))
                    {
                        _rows[rowIndex] = new RowIndexer(this);
                        e = _rows[rowIndex].WaitTillAvaliable(cellIndex);
                    }
                    else
                    {
                        if (_rows[rowIndex].TryAndGetCell(cellIndex, out ICellData d))
                        {
                            cell = d;
                            return;
                        }
                        else
                        {
                            e = _rows[rowIndex].WaitTillAvaliable(cellIndex);
                        }
                    }

                }
                if (e != null&& e.WaitOne(timeout))
                {
                    using ( await _lockRow.LockAsync())
                    {
                        if (_rows[rowIndex].TryAndGetCell(cellIndex, out ICellData d))
                        {
                            cell = d;
                            return;
                        }
                    }
                }

            });
            return cell;
        }


        /// <summary>
        /// This is for the ExcelImporter to add cells after they have been parsed
        /// </summary>
        /// <param name="cellData"></param>
        public async Task Add(ICellData cellData)
        {
            /*Task t1 = Task.Run(async() =>
            {*/
            using(await _lockRow.LockAsync())
            {
                if (!_rows.ContainsKey(cellData.CellRowIndex))
                {
                    _rows[cellData.CellRowIndex] = new RowIndexer(this);
                }

                _rows[cellData.CellRowIndex].Add(cellData);
                if(_listeners != null)
                {
                    //intresting so this will call each method and won't await till it's finished calling all of them
                    //pretty handy neat pattern
                    await Task.WhenAll(_listeners.Select(x => x.NotifyAsync(cellData)));
                }
            }
            //});
            /*Task t2 = Task.Run(async() =>
            {*/
            using (await _lockColumn.LockAsync())
            {
                if (!_columns.ContainsKey(cellData.CellColumnIndex))
                {
                    _columns[cellData.CellColumnIndex] = new ColumnIndexer(this);
                }

                _columns[cellData.CellColumnIndex].Add(cellData);
            }
            //});

            //await t1;
            //await t2;
        }

        /// <summary>
        /// Excel Import is finished loading we need to notify anything that is wait that today is not there day 
        /// and there not getting any data
        /// </summary>
        public void FinishedLoading()
        {
            _allLoaded = true;
            foreach(KeyValuePair<uint, RowIndexer> terminateWait in _rows)
            {
                terminateWait.Value.ClearNotify();
            }
            if(_listeners != null)
            {
                foreach (IItemEnquedEvent i in _listeners)
                {
                    i.FinishedLoading();
                }
            }
        }

        public bool AllLoaded => _allLoaded;

        AsyncLock ISpreadSheetGridCollectionAccessor.RowLock => _lockRow;

        Dictionary<uint, RowIndexer> ISpreadSheetGridCollectionAccessor.Rows => _rows;

        void ICollectionAccessor.EnqueListener(IItemEnquedEvent listener)
        {
            if (_listeners == null)
            {
                _listeners = new List<IItemEnquedEvent>();
            }
            _listeners.Add(listener);//this has the rowlock so it's thread safe
        }

        public IAsyncEnumerator<ICellData> GetAsyncEnumerator(CancellationToken cancellationToken = default) => new CellEnumerator(this);
        public IAsyncEnumerator<ICellData> GetAsyncEnumerator(IConditions conditions,CancellationToken cancellationToken = default) => new CellEnumerator(this,conditions);
    }
}
