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
        uint GetAllRows();
        bool TryGetRow(uint desiredRowIndex, out IEnumerator<Cell> cellEnumerator);
        ICellData ProcessedCell(Cell cellElement, ICellIndex index);
    }
}
