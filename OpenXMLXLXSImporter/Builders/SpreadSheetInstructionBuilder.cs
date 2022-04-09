using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
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
        public ISpreadSheetInstructionKey LoadColumnRange(uint row, string startColumn, string endColumn)
            => Add(new ColumnRange(row, startColumn, endColumn));

        public ISpreadSheetInstructionKey LoadRowRange(string column, uint startRow, uint endRow)
            => Add(new RowRange(column, startRow, endRow));

        public ISpreadSheetInstructionKey LoadFullColumnRange(uint row, string startColumn = "A")
            => Add(new FullColumnRange(row, startColumn));

        public ISpreadSheetInstructionKey LoadFullRowRange(string column, uint startRow)
            => Add(new FullRowRange(column, startRow));

        public async Task ProcessInstructions(SpreadSheetInstructionManager grid)
        {
            foreach(KeyValuePair<ISpreadSheetInstructionKey, ISpreadSheetInstruction> instruction in _instructions)
                await grid.ProcessInstruction(instruction.Value);
        }

        public async Task ProcessResults()
        {
            foreach (KeyValuePair<ISpreadSheetInstructionKey,ISpreadSheetInstruction> kv in _instructions)
            {
                if (kv.Key is SpreadSheetActionManager ssam)
                {
                    await ssam.TriggerEvent(kv.Value);
                }
            }

        }

        public async IAsyncEnumerable<ICellData> GetProcessedResults(ISpreadSheetInstructionKey key)
        {
            IAsyncEnumerable<ICellData>  cellDatas = _instructions[key].GetResults();

            IAsyncEnumerator<ICellData> cdEnum = cellDatas.GetAsyncEnumerator();

            while(await cdEnum.MoveNextAsync())
            {
                yield return cdEnum.Current;
            }
        }
    }
}
