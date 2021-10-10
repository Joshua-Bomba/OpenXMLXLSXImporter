using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.CellData
{
    public interface ICellData
    {
        string CellColumnIndex { get; set; }
        uint CellRowIndex { get; set; }
        string Content();
    }
}
