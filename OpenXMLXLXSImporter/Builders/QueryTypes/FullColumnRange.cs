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
        ColumnRange _internalRange;
        public FullColumnRange(uint row,string startingColumn = "A")
        {
            _row = row;
            _startingColumn = startingColumn;
            _internalRange = null;
        }

        public bool ByColumn => true;

        void ISpreadSheetInstruction.EnqueCell(IIndexer indexer)
        {
            string column = indexer.ColumnLength(_row);
            throw new NotImplementedException();
        }

        Task<IEnumerable<Task<ICellData>>> ISpreadSheetInstruction.GetResults()
        {
            throw new NotImplementedException();
        }
    }
}
