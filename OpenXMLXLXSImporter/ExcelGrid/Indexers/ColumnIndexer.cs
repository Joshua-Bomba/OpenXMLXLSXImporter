using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
{
    /// <summary>
    /// this manages access by cell then row
    /// </summary>
    public class ColumnIndexer : BaseSpreadSheetIndexer
    {
        private Dictionary<uint, ICellData> _cellsByRow;
        public ColumnIndexer(SpreadSheetGrid grid) : base(grid)
        {
            _cellsByRow = new Dictionary<uint, ICellData>();
        }

        public override void Add(ICellData cellData)
        {
            _cellsByRow[cellData.CellRowIndex] = cellData;
        }

    }
}
