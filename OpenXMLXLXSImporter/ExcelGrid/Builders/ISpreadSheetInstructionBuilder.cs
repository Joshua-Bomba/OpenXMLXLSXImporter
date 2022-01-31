using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Builders
{
    public  interface ISpreadSheetInstructionBuilder
    {
        ISpreadSheetInstruction LoadSingleCell(uint row, string cell);
    }
}
