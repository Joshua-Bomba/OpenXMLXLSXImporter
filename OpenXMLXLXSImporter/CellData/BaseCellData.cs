using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.CellData
{
    public abstract class BaseCellData : ICellData
    {
        public string CellColumnIndex { get; set; }
        public uint CellRowIndex { get; set; }

        public abstract string Content();
    }

}
