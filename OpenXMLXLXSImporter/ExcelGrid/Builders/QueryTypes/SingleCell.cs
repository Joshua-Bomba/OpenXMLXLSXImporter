using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.ExcelGrid.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Builders
{
    public class SingleCell : BaseSpreadSheetInstruction
    {
        private uint _row;
        private string _cell;
        private ICellIndex _cellItem;
        public SingleCell(uint row, string cellIndex) : base()
        {
            _row = row;
            _cell = cellIndex;
            _cellItem = null;
        }

        protected override IEnumerable<ICellIndex> GetResults() => new ICellIndex[] { _cellItem };

        protected override void EnqueCell(IIndexer indexer)
        {
            if(! indexer.HasCell(_row, _cell))
            {
                FutureCell item = new FutureCell(_row, _cell);
                indexer.Add(item);
                _cellItem = item;
            }
            else
            {
                _cellItem = indexer.GetCell(_row, _cell);
            }
            
        }


    }
}
