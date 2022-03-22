using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
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
        private IIndexer _indexer;
        public FullColumnRange(uint row,string startingColumn = "A")
        {
            _row = row;
            _startingColumn = startingColumn;
        }

        public bool IndexedByRow => false;

        void ISpreadSheetInstruction.EnqueCell(IIndexer indexer)
        {
            _indexer = indexer;
            _lastColumn = new LastColumn(_row);
            indexer.Add(_lastColumn);
        }

        async IAsyncEnumerable<ICellData> ISpreadSheetInstruction.GetResults()
        {
            ICellData lastCell = await _lastColumn.GetData();

            uint startingColumn = ExcelColumnHelper.GetColumnStringAsIndex(_startingColumn);
            uint lastColumn = ExcelColumnHelper.GetColumnStringAsIndex(lastCell.CellColumnIndex);
            lastColumn--;
            ISpreadSheetInstruction cr = new ColumnRange(lastCell.CellRowIndex, startingColumn, lastColumn);
            await _indexer.ProcessInstruction(cr);

            IAsyncEnumerable<ICellData> result = cr.GetResults();
            IAsyncEnumerator<ICellData> resultEnumerator = result.GetAsyncEnumerator();

            while(await resultEnumerator.MoveNextAsync())
            {
                yield return resultEnumerator.Current;
            }

            yield return await BaseSpreadSheetInstruction.GetCellData(lastCell);
        }
    }
}
