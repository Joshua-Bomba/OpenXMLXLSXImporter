using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.FileAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData.CellParsing
{
    public interface ICellParser
    {
        void AttachFileAccess(IXlsxDocumentFile file);
        ICellData ProcessCell(Cell cellElement, ICellIndex index);
    }
}
