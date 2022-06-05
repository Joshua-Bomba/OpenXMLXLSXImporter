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
        IAsyncEnumerable<Row> GetRows(string sheetName);

        Task<CellFormat> GetCellFormat(int index);

        Task<Fill> GetCellFill(CellFormat cellFormat);

        Task<DocumentFormat.OpenXml.Drawing.Color2Type> GetColorByThemeIndex(uint index);

        Task<OpenXmlElement> GetSharedStringTableElement(int index);
    }
}
