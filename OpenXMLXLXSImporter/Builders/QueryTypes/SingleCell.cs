using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public class SingleCell : BaseSpreadSheetInstruction
    {
        private uint _row;
        private string _cell;
        private Task<ICellIndex> _cellItem;

        public SingleCell(uint row, string cellIndex) : base()
        {
            _row = row;
            _cell = cellIndex;
            _cellItem = null;
        }

        protected async override IAsyncEnumerable<ICellIndex> GetResults()
        {
            yield return await _cellItem;
        }

        protected override void EnqueCell(IDataStore indexer)
        {
            _cellItem = indexer.GetCell(_row, _cell);
        }
    }
}
