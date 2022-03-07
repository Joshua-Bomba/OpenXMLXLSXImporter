using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public class RowRange : BaseSpreadSheetInstruction
    {
        private uint _startRow;
        private uint _endRow;
        private string _column;
        private ICellIndex[] _cellItems;

        public RowRange(string column,uint startRow,uint endRow)
        {
            _startRow = startRow;
            _endRow = endRow;
            _column = column;
            _cellItems = null;
            ValidateRange();
            _cellItems = new ICellIndex[(_endRow - _startRow) + 1];
        }

        private void ValidateRange()
        {
            if(_startRow > _endRow)
            {
                //🤷 Guess I could just flip it but meh
                throw new ArgumentException("The End Row has to be greater then the Start Row");
            }
        }
        protected override void EnqueCell(IIndexer indexer)
        {
            for(uint i = _startRow;i <= _endRow;i++)
            {
                if(!indexer.HasCell(i,_column))
                {
                    FutureCell item = new FutureCell(i, _column);
                    indexer.Add(item);
                    _cellItems[i - _startRow] = item;
                }
                else
                {
                    _cellItems[i - _startRow] = indexer.GetCell(i, _column);
                }

            }
        }

        protected override IEnumerable<ICellIndex> GetResults() => _cellItems;
    }
}
