using OpenXMLXLSXImporter.FileAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellParsing
{
    public class CellParserFactory : ICellParserFactory
    {
        private IXlsxDocumentFile _file;

        public void AttachFileAccess(IXlsxDocumentFile file) => _file=file;

        public ICellParser CreateNewCellParse() => new DefaultCellParser(_file);
    }
}
