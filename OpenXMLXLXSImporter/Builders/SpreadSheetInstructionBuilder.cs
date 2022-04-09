using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.FileAccess;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nito.AsyncEx;

namespace OpenXMLXLSXImporter.Builders
{
    public class SpreadSheetInstructionBuilder : ISpreadSheetInstructionRunner, ISpreadSheetInstructionBundler, ISpreadSheetInstructionBuilder
    {

        private SpreadSheetInstructionManager ssim;

        public ISpreadSheetInstructionRunner Runner => this;

        public ISpreadSheetInstructionBundler Bundler => this;

        public SpreadSheetInstructionBuilder(SpreadSheetInstructionManager spreadSheetInstructionManager)
        {
            ssim = spreadSheetInstructionManager;
        }

        public ISpreadSheetInstruction LoadSingleCell(uint row, string cell)
            => new SingleCell(row, cell);
        public ISpreadSheetInstruction LoadColumnRange(uint row, string startColumn, string endColumn)
            => new ColumnRange(row, startColumn, endColumn);

        public ISpreadSheetInstruction LoadRowRange(string column, uint startRow, uint endRow)
            => new RowRange(column, startRow, endRow);

        public ISpreadSheetInstruction LoadFullColumnRange(uint row, string startColumn)
            => new FullColumnRange(row, startColumn);

        public ISpreadSheetInstruction LoadFullRowRange(string column, uint startRow)
            => new FullRowRange(column, startRow);


        public async Task BundleRequeset(IEnumerable<ISpreadSheetInstruction> sheetInstructions)
        {
            await ssim.ProcessInstructionBundle(sheetInstructions);
        }

        public async IAsyncEnumerable<ICellData> GetBundledResults(IEnumerable<ISpreadSheetInstruction> ins)
        {
            ISpreadSheetInstruction[] instructions = ins.ToArray();
            Task t = ssim.ProcessInstructionBundle(instructions);
            IDictionary<ISpreadSheetInstruction, IAsyncEnumerable<ICellData>> dic = new Dictionary<ISpreadSheetInstruction, IAsyncEnumerable<ICellData>>(instructions.Length);
            IAsyncEnumerator<ICellData> enumerator = null;
            await t;
            foreach(ISpreadSheetInstruction instruction in instructions)
            {
                enumerator = instruction.GetResults().GetAsyncEnumerator();
                while(await enumerator.MoveNextAsync())
                {
                    yield return enumerator.Current;
                }
            }
        }

        async IAsyncEnumerable<ICellData> ISpreadSheetInstructionRunner.LoadCustomInstruction(ISpreadSheetInstruction instruction)
        {
            await ssim.ProcessInstruction(instruction);
            IAsyncEnumerator<ICellData> en =  instruction.GetResults().GetAsyncEnumerator();
            while(await en.MoveNextAsync())
            {
                yield return en.Current;
            }
        }

        async Task<ICellData> ISpreadSheetInstructionRunner.LoadSingleCell(uint row, string cell)
        {
            ISpreadSheetInstruction instr = LoadSingleCell(row,cell);
            await ssim.ProcessInstruction(instr);
            return await instr.GetResults().FirstOrDefaultAsync();
        }

        IAsyncEnumerable<ICellData> ISpreadSheetInstructionRunner.LoadColumnRange(uint row, string startColumn, string endColumn)
            => Runner.LoadCustomInstruction(this.LoadColumnRange(row, startColumn, endColumn));

        IAsyncEnumerable<ICellData> ISpreadSheetInstructionRunner.LoadRowRange(string column, uint startRow, uint endRow)
            => Runner.LoadCustomInstruction(this.LoadRowRange(column,startRow, endRow));

        IAsyncEnumerable<ICellData> ISpreadSheetInstructionRunner.LoadFullColumnRange(uint row, string startColumn)
            => Runner.LoadCustomInstruction(this.LoadFullColumnRange(row, startColumn));

        IAsyncEnumerable<ICellData> ISpreadSheetInstructionRunner.LoadFullRowRange(string column, uint startRow)
            => Runner.LoadCustomInstruction(this.LoadFullRowRange(column, startRow));

    }
}
