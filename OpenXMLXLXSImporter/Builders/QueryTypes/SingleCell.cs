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

        public override bool IndexedByRow => true;

        public SingleCell(uint row, string cellIndex) : base()
        {
            _row = row;
            _cell = cellIndex;
            _cellItem = null;
        }

        protected override IEnumerable<ICellIndex> GetResults() => new ICellIndex[] { _cellItem };

        protected override void EnqueCell(IIndexer indexer)
        {
            if(! indexer.TryGetCell(_row, _cell,out ICellIndex ci))
            {
                FutureCell item = new FutureCell(_row, _cell);
                indexer.Add(item);
                _cellItem = item;
            }
            else
            {
                _cellItem = ci;
            }
        }
    }
}
