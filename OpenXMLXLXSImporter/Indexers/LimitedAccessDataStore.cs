using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public class LimitedAccessDataStore : IDataStoreLocked
    {
        private DirectDataStore _d;
        public LimitedAccessDataStore(DirectDataStore d)
        {
            _d = d;
        }
        public void Delete()
        {
            _d = null;
        }

        public ICellIndex GetCell(uint rowIndex, string cellIndex) => _d.GetCell(rowIndex, cellIndex);

        public LastColumn GetLastColumn(uint rowIndex) => _d.GetLastColumn(rowIndex);

        public LastRow GetLastRow() => _d.GetLastRow();
    }
}
