using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    /// <summary>
    /// this manages access by cell then row
    /// </summary>
    public class ColumnIndexer : BaseIndexer
    {
        private Dictionary<string,Dictionary<uint, ICellIndex>> _cells;
        public ColumnIndexer(ISpreadSheetIndexersLock indexerLock) : base(indexerLock)
        {
            _cells = new Dictionary<string, Dictionary<uint, ICellIndex>>();
        }

        protected override void InternalSet(ICellIndex cell)
        {
            if (!_cells.ContainsKey(cell.CellColumnIndex))
            {
                _cells.Add(cell.CellColumnIndex, new Dictionary<uint, ICellIndex>());
            }
            _cells[cell.CellColumnIndex][cell.CellRowIndex] =  cell;
        }

        protected override void InternalAdd(ICellIndex cell)
        {
            if(!_cells.ContainsKey(cell.CellColumnIndex))
            {
                _cells.Add(cell.CellColumnIndex, new Dictionary<uint, ICellIndex>());
            }
            _cells[cell.CellColumnIndex].Add(cell.CellRowIndex, cell);
        }

        public override bool TryGetCell(uint rowIndex, string cellIndex, out ICellIndex ci)
        {
            if(_cells.ContainsKey(cellIndex)&&_cells[cellIndex].ContainsKey(rowIndex))
            {
                ci = _cells[cellIndex][rowIndex];
                return true;
            }
            else
            {
                ci = null;
                return false;
            }
        }

        //public override bool HasCell(uint rowIndex, string cellIndex)
        //    => _cells.ContainsKey(cellIndex)&&_cells[cellIndex].ContainsKey(rowIndex);

        //public override ICellIndex GetCell(uint rowIndex, string cellIndex)
        //    => _cells[cellIndex][rowIndex];

    }
}
