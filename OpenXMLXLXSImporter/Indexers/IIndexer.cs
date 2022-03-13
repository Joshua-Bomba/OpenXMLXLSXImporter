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

        uint RowLength(string cellIndex);

        string ColumnLength(uint rowIndex);

        void Spread(ICellIndex cell);
    }
}
