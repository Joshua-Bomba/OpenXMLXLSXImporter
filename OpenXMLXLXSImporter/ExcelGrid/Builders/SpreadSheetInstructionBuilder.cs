using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Builders
{
    public class SpreadSheetInstructionBuilder : ISpreadSheetInstructionBuilder, ISpreadSheetQueryResults
    {
        private Dictionary<ISpreadSheetInstructionKey, ISpreadSheetInstruction> _instructions;
        public SpreadSheetInstructionBuilder()
        {
            _instructions = new Dictionary<ISpreadSheetInstructionKey, ISpreadSheetInstruction>();
        }


        protected virtual ISpreadSheetInstructionKey Add(ISpreadSheetInstruction i)
        {
            SpreadSheetActionManager ssam = new SpreadSheetActionManager();
            _instructions.Add(ssam, i);
            return ssam;
        }

        public ISpreadSheetInstructionKey LoadSingleCell(uint row, string cell)
            => Add(new SingleCell(row, cell));

        public async Task ProcessInstructions(SpreadSheetGrid grid)
        {
            foreach(KeyValuePair<ISpreadSheetInstructionKey, ISpreadSheetInstruction> instruction in _instructions)
                await grid.ProcessInstruction(instruction.Value);
        }

        public async Task ProcessResults()
        {
            foreach (Task t in _instructions.Select(x => (x.Key as SpreadSheetActionManager)?.TriggerEvent(x.Value)).ToArray())
                await t;
        }

        public async Task<IEnumerable<Task<ICellData>>> GetResults(ISpreadSheetInstructionKey key)
            =>await _instructions[key].GetResults();
    }
}
