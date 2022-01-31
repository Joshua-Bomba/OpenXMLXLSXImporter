using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Builders
{
    public interface ISpreadSheetInstructionKey
    {
        delegate void SpreadSheetAction(IEnumerable<Task<ICellData>> ac);
        void AddAction(SpreadSheetAction data);
    }
}
