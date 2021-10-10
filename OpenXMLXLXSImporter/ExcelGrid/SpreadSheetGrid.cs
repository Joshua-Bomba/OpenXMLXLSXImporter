using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.ExcelGrid.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid
{
    public class SpreadSheetGrid
    {
        ISheetProperties _sheet;
        private Dictionary<uint, RowIndexer> _rows;
        private Dictionary<string, ColumnIndexer> _columns;
        private ManualResetEvent _allLoaded;
        private object _lockRow = new object();
        private object _lockColumn = new object();
        public SpreadSheetGrid(ISheetProperties sheet)
        {
            _sheet = sheet;
            sheet.SetCellMatrix(this);
            _allLoaded = new ManualResetEvent(false);
            _rows = new Dictionary<uint, RowIndexer>();
            _columns = new Dictionary<string, ColumnIndexer>();
        }

        public bool TryFetchCell(uint row, string cell)
        {
            throw new NotImplementedException();
        }

        public async Task<ICellData> FetchCell(uint rowIndex, string cellIndex, int timeout = -1)
        {
            ICellData cell = null;
            await Task.Run(() =>
            {
                ManualResetEvent e = null;
                lock (_lockRow)
                {
                    if (!_rows.ContainsKey(rowIndex))
                    {
                        _rows[rowIndex] = new RowIndexer();
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
                if (e.WaitOne(timeout))
                {
                    lock (_lockRow)
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

        public void Add(ICellData cellData)
        {
            Task t1 = Task.Run(() =>
            {
                lock (_lockRow)
                {
                    if (!_rows.ContainsKey(cellData.CellRowIndex))
                    {
                        _rows[cellData.CellRowIndex] = new RowIndexer();
                    }

                    _rows[cellData.CellRowIndex].Add(cellData);
                }
            });
            lock (_lockColumn)
            {
                if (!_columns.ContainsKey(cellData.CellColumnIndex))
                {
                    _columns[cellData.CellColumnIndex] = new ColumnIndexer();
                }

                _columns[cellData.CellColumnIndex].Add(cellData);
            }

            t1.Wait();
        }

        public void FinishedLoading()
        {
            _allLoaded.Set();
        }
    }
}
