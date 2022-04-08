using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
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
        private IDataStore _indexer;
        public FullRowRange(string column, uint startRow)
        {
            _column = column;
            _startRow = startRow;
        }

        public bool IndexedByRow => true;

        async Task ISpreadSheetInstruction.EnqueCell(IDataStore indexer)
        {
            _indexer = indexer;
            _lastRow = await _indexer.GetLastRow();
        }

        async IAsyncEnumerable<ICellData> ISpreadSheetInstruction.GetResults()
        {
            ICellData lastCell = await _lastRow.GetData();

            ISpreadSheetInstruction rr = new RowRange(_column, _startRow, lastCell.CellRowIndex - 1);
            await _indexer.ProcessInstruction(rr);

            IAsyncEnumerable<ICellData> result = rr.GetResults();
            IAsyncEnumerator<ICellData> resultEnumerator = result.GetAsyncEnumerator();

            while(await resultEnumerator.MoveNextAsync())
            {
                yield return resultEnumerator.Current;
            }
            yield return await BaseSpreadSheetInstruction.GetCellData(lastCell);

        }
    }
}
