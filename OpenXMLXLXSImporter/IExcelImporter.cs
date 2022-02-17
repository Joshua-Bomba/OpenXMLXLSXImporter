using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter
{
    public interface IExcelImporter : IDisposable
    {
        Task Process(ISheetProperties prop);
    }
}
