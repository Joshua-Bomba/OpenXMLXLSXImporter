using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.FileAccess
{
    public interface IXlsxDocumentFile
    {
        Sheet GetSheet(string sheetName);
        WorkbookPart WorkbookPart { get; }

        CellFormat GetCellFormat(int index);

        OpenXmlElement GetSharedStringTableElement(int index);
    }
}
