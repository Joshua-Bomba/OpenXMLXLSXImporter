using Nito.AsyncEx;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    /// <summary>
    /// this manages access by row then cell
    /// </summary>
    public class RowIndexer 
    {
        //protected Dictionary<uint,Dictionary<string, ICellIndex>> _cells;
        protected Dictionary<uint, ColumnIndexer> _cells;
        protected LastRow _lastRow;
        public RowIndexer()
        {
            _cells = new Dictionary<uint, ColumnIndexer>();
            //_cells = new Dictionary<uint, Dictionary<string, ICellIndex>>();
            //_lastRow = null;
        }

        public  ICellIndex Get(uint rowIndex, string cellIndex)
        {
            if (_cells.ContainsKey(rowIndex) && _cells[rowIndex].ContainsKey(cellIndex))
            {
                return _cells[rowIndex][cellIndex];
            }
            return null;
        }

        public void Set(ICellIndex cellData)
        {
            if (!_cells.ContainsKey(cellData.CellRowIndex))
            {
                _cells.Add(cellData.CellRowIndex, new ColumnIndexer());
            }
            _cells[cellData.CellRowIndex][cellData.CellColumnIndex] = cellData;
        }

        public void Add(ICellIndex cellData)
        {
            if (!_cells.ContainsKey(cellData.CellRowIndex))
            {
                _cells.Add(cellData.CellRowIndex, new ColumnIndexer());
            }
            _cells[cellData.CellRowIndex].Add(cellData.CellColumnIndex, cellData);
        }
    }
}
