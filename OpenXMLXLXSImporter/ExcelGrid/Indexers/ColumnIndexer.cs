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
    public class ColumnIndexer : BaseIndexer
    {
        private Dictionary<string,Dictionary<uint, ICellIndex>> _cells;
        public ColumnIndexer() : base()
        {
            _cells = new Dictionary<string, Dictionary<uint, ICellIndex>>();
        }

        public async override Task Add(ICellIndex cellData)
        {
            //_cellsByRow[cellData.CellRowIndex] = cellData;
        }

        public async override Task<ICellData> GetCell(uint rowIndex, string cellIndex)
        {
            throw new NotImplementedException();
        }

        public async override Task<bool> HasCell(uint rowIndex, string cellIndex)
        {
            throw new NotImplementedException();
        }
    }
}
