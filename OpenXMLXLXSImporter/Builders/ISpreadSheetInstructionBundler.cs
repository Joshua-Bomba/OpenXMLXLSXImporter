using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public interface ISpreadSheetInstructionBundler
    {
        ISpreadSheetInstruction LoadSingleCell(uint row, string cell);

        ISpreadSheetInstruction LoadColumnRange(uint row, string startColumn, string endColumn);

        ISpreadSheetInstruction LoadRowRange(string column, uint startRow, uint endRow);

        ISpreadSheetInstruction LoadFullColumnRange(uint row, string startColumn = "A");

        ISpreadSheetInstruction LoadFullRowRange(string column, uint startRow = 1);

        Task BundleRequeset();

        IAsyncEnumerable<ICellData> GetBundledResults();

    }
}
