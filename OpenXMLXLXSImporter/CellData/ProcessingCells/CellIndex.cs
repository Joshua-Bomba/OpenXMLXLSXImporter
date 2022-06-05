using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public  class CellIndex : ICellIndex
    {
        public string CellColumnIndex { get; set; }
        public uint CellRowIndex { get; set; }
    }
}
