using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using OpenXMLXLSXImporter.Processing;
using OpenXMLXLSXImporter.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public class FullColumnRange : ISpreadSheetInstruction
    {
        private uint _row;
        private string _startingColumn;
        private LastColumn _lastColumn;
        private ISpreadSheetInstructionManager _manger;
        public FullColumnRange(uint row,string startingColumn = "A")
        {
            _row = row;
            _startingColumn = startingColumn;
        }

        public void AttachSpreadSheetInstructionManager(ISpreadSheetInstructionManager spreadSheetInstructionManager)
        {
            _manger = spreadSheetInstructionManager;
        }

        void ISpreadSheetInstruction.EnqueCell(IDataStoreLocked indexer)
        {
            _lastColumn = indexer.GetLastColumn(_row);
        }

        async IAsyncEnumerable<ICellData> ISpreadSheetInstruction.GetResults()
        {
            ICellIndex lastCell = await _lastColumn.GetIndex();
            if(lastCell != null)
            {
                uint startingColumn = ExcelColumnHelper.GetColumnStringAsIndex(_startingColumn);
                uint lastColumn = ExcelColumnHelper.GetColumnStringAsIndex(lastCell.CellColumnIndex);
                ISpreadSheetInstruction cr = new ColumnRange(lastCell.CellRowIndex, startingColumn, lastColumn);
                await _manger.ProcessInstruction(cr);

                IAsyncEnumerable<ICellData> result = cr.GetResults();
                IAsyncEnumerator<ICellData> resultEnumerator = result.GetAsyncEnumerator();

                while (await resultEnumerator.MoveNextAsync())
                {
                    yield return resultEnumerator.Current;
                }
            }
        }
    }
}
