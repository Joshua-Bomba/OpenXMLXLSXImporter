using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.ExcelGrid.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Builders
{
    public interface ISpreadSheetInstruction
    {
        bool ByColumn { get; }

        void EnqueCell(IIndexer indexer);

        Task<IEnumerable<Task<ICellData>>> GetResults();
    }
}
