using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public interface ISpreadSheetInstructionKey
    {
        delegate void SpreadSheetAction(IEnumerable<Task<ICellData>> ac);
        void AddAction(SpreadSheetAction data);
    }
}
