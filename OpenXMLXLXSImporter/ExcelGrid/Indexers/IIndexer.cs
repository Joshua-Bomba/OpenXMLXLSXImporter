using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
{
    public interface IIndexer
    {
        Task Add(ICellData cellData);

        Task<bool>  HasCell(uint rowIndex, string cellIndex);

        Task<ICellData> GetCell(uint rowIndex, string cellIndex);
    }
}
