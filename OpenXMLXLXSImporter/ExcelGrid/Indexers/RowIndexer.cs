using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
{
    /// <summary>
    /// this manages access by row then cell
    /// </summary>
    public class RowIndexer : BaseSpreadSheetIndexer
    {
        protected Dictionary<string, ICellData> _cellsByColumn;
        private Dictionary<string, ManualResetEvent> _notify;
        public RowIndexer(SpreadSheetGrid grid) : base(grid)
        {
            _cellsByColumn = new Dictionary<string, ICellData>();
            _notify = null;
        }
        /// <summary>
        /// this will add a cell called by the ExcelImporter
        /// </summary>
        /// <param name="cellData"></param>
        public override void Add(ICellData cellData)
        {
            _cellsByColumn.Add(cellData.CellColumnIndex, cellData);
            if (_notify != null && _notify.ContainsKey(cellData.CellColumnIndex))
            {
                _notify[cellData.CellColumnIndex].Set();
            }
        }
        /// <summary>
        /// this will get a cell if it's avaliable
        /// </summary>
        /// <param name="cellIndex"></param>
        /// <param name="d"></param>
        /// <returns></returns>
        public bool TryAndGetCell(string cellIndex, out ICellData d)
        {
            bool containsCell = _cellsByColumn.ContainsKey(cellIndex);
            if (containsCell)
            {
                d = _cellsByColumn[cellIndex];
            }
            else
            {
                d = null;
            }
            return containsCell;
        }
        /// <summary>
        /// Gets the WaitTillAvaliable ManualResetEvent which can be waited on
        /// </summary>
        /// <param name="cellIndex"></param>
        /// <returns></returns>
        public ManualResetEvent WaitTillAvaliable(string cellIndex)
        {
            if(!_grid.AllLoaded)
            {
                if (_notify != null)
                {
                    _notify = new Dictionary<string, ManualResetEvent>();
                }
                if (!_notify.ContainsKey(cellIndex))
                {
                    _notify.Add(cellIndex, new ManualResetEvent(false));
                }
                return _notify[cellIndex];
            }
            return null;
        }


        public void ClearNotify()
        {
            if (_notify != null)
            {
                foreach(KeyValuePair<string,ManualResetEvent> kv in _notify)
                {
                    kv.Value.Set();
                }
            }
        }

    }
}
