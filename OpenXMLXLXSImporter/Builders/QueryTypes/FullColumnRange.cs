using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public class FullColumnRange : ISpreadSheetInstruction
    {
        private uint _row;
        private string _startingColumn;
        private Task<IEnumerable<ICellIndex>> d;
        public FullColumnRange(uint row,string startingColumn = "A")
        {
            _row = row;
            _startingColumn = startingColumn;
        }

        public bool ByColumn => true;

        void ISpreadSheetInstruction.EnqueCell(IIndexer indexer)
        {
            d = indexer.ToMaxColumnLength(_row,_startingColumn);
        }

        async IAsyncEnumerable<ICellData> ISpreadSheetInstruction.GetResults()
        {
            foreach(ICellIndex cell in await d)
            {
                yield return await BaseSpreadSheetInstruction.GetCellData(cell);
            }

        }
    }
}
