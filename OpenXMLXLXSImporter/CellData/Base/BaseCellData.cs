using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    /// <summary>
    /// an abstract class for ICellData so we don't have to bother implement common properties again
    /// </summary>
    public abstract class BaseCellData : ICellData
    {
        public string CellColumnIndex { get; set; }
        public uint CellRowIndex { get; set; }

        public abstract string Content();
    }

}
