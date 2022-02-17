using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public interface ISpreadSheetInstruction
    {
        bool ByColumn { get; }

        void EnqueCell(IIndexer indexer);

        Task<IEnumerable<Task<ICellData>>> GetResults();
    }
}
