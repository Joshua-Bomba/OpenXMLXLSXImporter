using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public class FullRowRange : BaseSpreadSheetInstruction
    {
        private string _column;
        private uint _startRow;
        private IEnumerable<ICellIndex> _results;
        public FullRowRange(string column, uint startRow)
        {
            _column = column;
            _startRow = startRow;
        }
        public override bool IndexedByRow => true;

        protected override void EnqueCell(IIndexer indexer)
        {
            //_results = indexer.GetFullColumnRows(_column, _startRow);
        }

        protected override IEnumerable<ICellIndex> GetResults() => _results;
    }
}
