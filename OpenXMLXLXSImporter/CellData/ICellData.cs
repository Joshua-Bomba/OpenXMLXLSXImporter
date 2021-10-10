using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.CellData
{
    public interface ICellData
    {
        /// <summary>
        /// this will hold the CellColumnIndex Example A, F, AA, ect
        /// </summary>
        string CellColumnIndex { get; set; }
        /// <summary>
        /// this will store the row #
        /// </summary>
        uint CellRowIndex { get; set; }
        /// <summary>
        /// basically a tostring for every type of cell example a date
        /// just incase you want a string
        /// </summary>
        /// <returns></returns>
        string Content();
    }
}
