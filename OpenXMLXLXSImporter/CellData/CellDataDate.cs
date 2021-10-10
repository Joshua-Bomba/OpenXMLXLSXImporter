using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.CellData
{
    /// <summary>
    /// A Cell that's data is a date
    /// </summary>
    public class CellDataDate : BaseCellData
    {
        public CellDataDate()
        {

        }
        public DateTime Date { get; set; }

        public string DateFormat { get; set; }

        public override string Content() => Date.ToString(DateFormat);
    }
}
