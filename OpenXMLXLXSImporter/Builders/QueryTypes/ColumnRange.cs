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
    public class ColumnRange : BaseSpreadSheetInstruction
    {
        private uint _row;
        private string _startColumn;
        private string _endColumn;

        private uint _startColumnInt;
        private uint _endColumnInt;

        private ICellIndex[] _cellItems;
        public ColumnRange(uint row, string startColumn, string endColumn)
        {
            _row = row;
            _startColumn = startColumn;
            _endColumn = endColumn;
            ValidateRange();
            _cellItems = new ICellIndex[_endColumnInt - _startColumnInt];
        }

        private void ValidateRange()
        {
            _startColumnInt = ExcelColumnHelper.GetColumnStringAsIndex(_startColumn);
            _endColumnInt = ExcelColumnHelper.GetColumnStringAsIndex(_endColumn);
            if(_startColumnInt > _endColumnInt)
            {//🤷 Guess I could just flip it but meh
                throw new ArgumentException("The End Column has to be greater then the Start Column");
            }
        }

        protected override void EnqueCell(IIndexer indexer)
        {
            for(uint i = _startColumnInt;i <= _endColumnInt;i++)
            {
                string column = ExcelColumnHelper.GetColumnIndexAsString(i);
                if(indexer.HasCell(_row,column))
                {
                    FutureCell item = new FutureCell(_row,column);
                    indexer.Add(item);
                    _cellItems[i - _startColumnInt] = item;
                }
                else
                {
                    _cellItems[i - _startColumnInt] = indexer.GetCell(_row,column);
                }
            }
        }

        protected override IEnumerable<ICellIndex> GetResults() => _cellItems;
    }
}
