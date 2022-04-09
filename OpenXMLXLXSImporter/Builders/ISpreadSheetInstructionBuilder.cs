using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public  interface ISpreadSheetInstructionBuilder
    {
        ISpreadSheetInstructionKey LoadSingleCell(uint row, string cell);

        ISpreadSheetInstructionKey LoadColumnRange(uint row, string startColumn, string endColumn);

        ISpreadSheetInstructionKey LoadRowRange(string column, uint startRow, uint endRow);

        ISpreadSheetInstructionKey LoadFullColumnRange(uint row, string startColumn = "A");

        ISpreadSheetInstructionKey LoadFullRowRange(string column, uint startRow = 1);
    }
}
