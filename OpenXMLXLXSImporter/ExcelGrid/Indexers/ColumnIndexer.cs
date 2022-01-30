using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
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

        protected override void InternalAdd(ICellIndex cell)
        {
            if(!_cells.ContainsKey(cell.CellColumnIndex))
            {
                _cells.Add(cell.CellColumnIndex, new Dictionary<uint, ICellIndex>());
            }
            _cells[cell.CellColumnIndex].Add(cell.CellRowIndex, cell);
        }

        protected override bool InternalContains(uint rowIndex, string cellIndex)
            => _cells.ContainsKey(cellIndex)&&_cells[cellIndex].ContainsKey(rowIndex);

        protected override ICellIndex InternalGet(uint rowIndex, string cellIndex)
            => _cells[cellIndex][rowIndex];
    }
}
