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
        private ICellIndex _cellItem;

        public SingleCell(uint row, string cellIndex) : base()
        {
            _row = row;
            _cell = cellIndex;
            _cellItem = null;
        }

        protected override IEnumerable<ICellIndex> GetResults()
        {
            yield return _cellItem;
        }

        protected override void EnqueCell(IDataStoreLocked indexer)
        {
            _cellItem = indexer.GetCell(_row, _cell);
        }
    }
}
