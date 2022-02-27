using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders.Managers
{
    public interface ISpreadSheet : IDisposable
    {
        Task<ICellData> LoadSingleCell(uint row, string cell);
    }
}
