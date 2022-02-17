using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.FileAccess
{
    public interface ISpreadSheetFile
    {
        Sheet GetSheet(string sheetName);
        WorkbookPart WorkbookPart { get; }
    }
}
