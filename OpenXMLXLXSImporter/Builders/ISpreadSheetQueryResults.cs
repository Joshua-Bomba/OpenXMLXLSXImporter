using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.Builders
{
    public interface ISpreadSheetQueryResults
    {
        Task<IEnumerable<Task<ICellData>>> GetResults(ISpreadSheetInstructionKey key);
    }
}
