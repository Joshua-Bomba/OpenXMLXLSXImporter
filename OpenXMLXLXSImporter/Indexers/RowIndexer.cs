using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.Indexers
{
    /// <summary>
    /// this manages access by row then cell
    /// </summary>
    public class RowIndexer : BaseIndexer
    {
        protected Dictionary<uint,Dictionary<string, ICellIndex>> _cells;
        public RowIndexer(ISpreadSheetIndexersLock indexerLock) : base(indexerLock)
        {
            _cells = new Dictionary<uint, Dictionary<string, ICellIndex>>();
        }

        protected override void InternalAdd(ICellIndex cellData)
        {
            if (!_cells.ContainsKey(cellData.CellRowIndex))
            {
                _cells.Add(cellData.CellRowIndex, new Dictionary<string, ICellIndex>());
            }
            _cells[cellData.CellRowIndex].Add(cellData.CellColumnIndex,cellData);
        }

        public override bool HasCell(uint rowIndex, string cellIndex)
            => _cells.ContainsKey(rowIndex)&&_cells[rowIndex].ContainsKey(cellIndex);


        public override ICellIndex GetCell(uint rowIndex, string cellIndex)
            => _cells[rowIndex][cellIndex];
    }
}
