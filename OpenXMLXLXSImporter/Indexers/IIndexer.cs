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

        void Add(ICellProcessingTask index);

        bool HasCell(uint rowIndex, string cellIndex);

        ICellIndex GetCell(uint rowIndex, string cellIndex);

        Task<IEnumerable<ICellIndex>> ToMaxRowLength(string cellIndex, int startRow);

        Task<IEnumerable<ICellIndex>> ToMaxColumnLength(uint rowIndex, string StartColumn);

        void Spread(ICellIndex cell);
    }
}
