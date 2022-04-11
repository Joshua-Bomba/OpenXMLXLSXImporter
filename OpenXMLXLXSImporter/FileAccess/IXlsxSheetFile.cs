using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.FileAccess
{
    public interface IXlsxSheetFile
    {
        Task<uint> GetAllRows();
        Task<IEnumerator<Cell>> GetRow(uint desiredRowIndex);
        ICellData ProcessedCell(Cell cellElement, ICellIndex index);
    }
}
