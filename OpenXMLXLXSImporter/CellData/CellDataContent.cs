using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    /// <summary>
    /// basic content of a cell
    /// </summary>
    public class CellDataContent : BaseCellData
    {
        public string Text { get; set; }
        public override string Content() => Text;
    }
}
