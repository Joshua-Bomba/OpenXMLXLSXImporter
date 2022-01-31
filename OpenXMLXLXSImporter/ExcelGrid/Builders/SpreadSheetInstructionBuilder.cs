using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Builders
{
    public class SpreadSheetInstructionBuilder : ISpreadSheetInstructionBuilder
    {
        private List<ISpreadSheetInstruction> _instructions;
        public SpreadSheetInstructionBuilder()
        {
            _instructions = new List<ISpreadSheetInstruction>();
        }
        public ISpreadSheetInstruction LoadSingleCell(uint row, string cell)
            => new SingleCell(row, cell);

        public async Task ProcessInstructions(SpreadSheetGrid grid)
        {
            foreach(ISpreadSheetInstruction instruction in _instructions)
                await grid.ProcessInstruction(instruction);
        }
    }
}
