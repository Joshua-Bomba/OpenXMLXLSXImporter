using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders.QueryTypes
{
    public class FullColumnRange : BaseSpreadSheetInstruction
    {
        private uint _row;
        private string _startingColumn;
        private IEnumerable<ICellIndex> _result;
        public FullColumnRange(uint row,string startingColumn = "A")
        {
            _row = row;
            _startingColumn = startingColumn;
        }
        protected override void EnqueCell(IIndexer indexer)
        {
            _result = indexer.GetFullRowColumns(_row, _startingColumn);
        }

        protected override IEnumerable<ICellIndex> GetResults() => _result;
    }
}
