using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.FileAccess
{
    public interface IXlsxDocumentFilePromise
    {
        Task<IXlsxDocumentFile> GetLoadedFile();
    }
}
