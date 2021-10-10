using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
{
    public class RowIndexer : BaseSpreadSheetIndexer
    {
        protected Dictionary<string, ICellData> _cellsByColumn;
        private Dictionary<string, ManualResetEvent> _notify;
        public RowIndexer() : base()
        {
            _cellsByColumn = new Dictionary<string, ICellData>();
            _notify = null;
        }

        public override void Add(ICellData cellData)
        {
            _cellsByColumn.Add(cellData.CellColumnIndex, cellData);
            if (_notify != null && _notify.ContainsKey(cellData.CellColumnIndex))
            {
                _notify[cellData.CellColumnIndex].Set();
            }
        }

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

        public ManualResetEvent WaitTillAvaliable(string cellIndex)
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

    }
}
