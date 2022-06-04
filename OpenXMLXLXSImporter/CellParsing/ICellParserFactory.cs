using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.FileAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellParsing
{
    public interface ICellParserFactory
    {
        void AttachFileAccess(IXlsxDocumentFile file);
        ICellParser CreateNewCellParse();
    }
}
