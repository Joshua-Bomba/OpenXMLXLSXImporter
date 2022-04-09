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
        IAsyncEnumerable<ICellData> LoadCustomInstruction(ISpreadSheetInstruction instruction);

        Task<ICellData> LoadSingleCell(uint row, string cell);

        IAsyncEnumerable<ICellData> LoadColumnRange(uint row, string startColumn, string endColumn);

        IAsyncEnumerable<ICellData> LoadRowRange(string column, uint startRow, uint endRow);

        IAsyncEnumerable<ICellData> LoadFullColumnRange(uint row, string startColumn = "A");

        IAsyncEnumerable<ICellData> LoadFullRowRange(string column, uint startRow = 1);
    }
}
