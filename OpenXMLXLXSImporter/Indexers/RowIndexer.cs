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
    public class RowIndexer : BaseIndexer
    {
        protected Dictionary<uint,Dictionary<string, ICellIndex>> _cells;
        public RowIndexer(ISpreadSheetIndexersLock indexerLock) : base(indexerLock)
        {
            _cells = new Dictionary<uint, Dictionary<string, ICellIndex>>();
        }

        public override bool TryGetCell(uint rowIndex, string cellIndex, out ICellIndex ci)
        {
            throw new NotImplementedException();
        }

        protected override void InternalAdd(ICellIndex cellData)
        {
            if (!_cells.ContainsKey(cellData.CellRowIndex))
            {
                _cells.Add(cellData.CellRowIndex, new Dictionary<string, ICellIndex>());
            }
            _cells[cellData.CellRowIndex].Add(cellData.CellColumnIndex,cellData);
        }
    }
}
