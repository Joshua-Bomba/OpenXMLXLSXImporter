using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public interface IDataStore
    {
        Task<ICellIndex> GetCell(uint rowIndex, string cellIndex);
        Task<LastColumn> GetLastColumn(uint rowIndex);
        Task<LastRow> GetLastRow();
    }
}
