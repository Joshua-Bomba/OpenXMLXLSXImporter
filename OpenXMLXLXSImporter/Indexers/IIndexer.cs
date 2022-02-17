using OpenXMLXLXSImporter.Builders;
using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.Indexers
{
    public interface IIndexer
    {
        Task ProcessInstruction(ISpreadSheetInstruction instruction);

        void Add(ICellProcessingTask index);

        bool  HasCell(uint rowIndex, string cellIndex);

        ICellIndex GetCell(uint rowIndex, string cellIndex);

        void Spread(ICellIndex cell);
    }
}
