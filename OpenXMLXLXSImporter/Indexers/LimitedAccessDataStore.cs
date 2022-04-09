using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public class LimitedAccessDataStore : IDataStore
    {
        private IDataStore _d;
        public LimitedAccessDataStore(IDataStore d)
        {
            _d = d;
        }
        public void Delete()
        {
            _d = null;
        }

        public Task<ICellIndex> GetCell(uint rowIndex, string cellIndex) => _d.GetCell(rowIndex, cellIndex);

        public Task<LastColumn> GetLastColumn(uint rowIndex) => _d.GetLastColumn(rowIndex);

        public Task<LastRow> GetLastRow() => _d.GetLastRow();
    }
}
