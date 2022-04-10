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
        Task<Sheet> GetSheet(string sheetName);
        Task<WorksheetPart> GetWorkSheetPartById(string sheetID);

        Task<CellFormat> GetCellFormat(int index);

        Task<OpenXmlElement> GetSharedStringTableElement(int index);
    }
}
