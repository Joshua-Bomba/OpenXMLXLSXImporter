using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public class FullRowRange : ISpreadSheetInstruction
    {
        private string _column;
        private uint _startRow;
        
        private LastRow _lastRow;
        private ISpreadSheetInstructionManager _manager;
        public FullRowRange(string column, uint startRow)
        {
            _column = column;
            _startRow = startRow;
        }

        public void AttachSpreadSheetInstructionManager(ISpreadSheetInstructionManager spreadSheetInstructionManager)
        {
            _manager = spreadSheetInstructionManager;
        }

        async Task ISpreadSheetInstruction.EnqueCell(IDataStore indexer)
        {
            _lastRow = await indexer.GetLastRow();
        }

        async IAsyncEnumerable<ICellData> ISpreadSheetInstruction.GetResults()
        {
            ICellIndex lastCell = await _lastRow.GetIndex();

            ISpreadSheetInstruction rr = new RowRange(_column, _startRow, lastCell.CellRowIndex);
            await _manager.ProcessInstruction(rr);

            IAsyncEnumerable<ICellData> result = rr.GetResults();
            IAsyncEnumerator<ICellData> resultEnumerator = result.GetAsyncEnumerator();

            while(await resultEnumerator.MoveNextAsync())
            {
                yield return resultEnumerator.Current;
            }
        }
    }
}
