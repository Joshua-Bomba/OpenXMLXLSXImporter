using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public interface ISpreadSheetInstructionRunner
    {
        IAsyncEnumerable<ICellIndex> LoadCustomInstruction(ISpreadSheetInstruction instruction);

        Task<ICellData> LoadSingleCell(uint row, string cell);

        IAsyncEnumerable<ICellIndex> LoadColumnRange(uint row, string startColumn, string endColumn);

        IAsyncEnumerable<ICellIndex> LoadRowRange(string column, uint startRow, uint endRow);

        IAsyncEnumerable<ICellIndex> LoadFullColumnRange(uint row, string startColumn = "A");

        IAsyncEnumerable<ICellIndex> LoadFullRowRange(string column, uint startRow = 1);
    }
}
