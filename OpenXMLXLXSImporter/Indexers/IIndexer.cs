using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public interface IIndexer
    {
        Task ProcessInstruction(ISpreadSheetInstruction instruction);

        Task Add(ICellProcessingTask index);

        bool TryGetCell(uint rowIndex, string cellIndex, out ICellIndex ci);

        Task SetCell(ICellData d);

        void Spread(ICellIndex cell);
    }
}
