using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders.QueryTypes
{
    public class FullRowRange : BaseSpreadSheetInstruction
    {
        private string _column;
        public FullRowRange(string column)
        {

        }
        protected override void EnqueCell(IIndexer indexer)
        {
            throw new NotImplementedException();
        }

        protected override IEnumerable<ICellIndex> GetResults()
        {
            throw new NotImplementedException();
        }
    }
}
