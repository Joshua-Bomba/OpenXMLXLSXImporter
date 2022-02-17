using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public interface ICellData : ICellIndex
    {
        /// <summary>
        /// basically a tostring for every type of cell example a date
        /// just incase you want a string
        /// </summary>
        /// <returns></returns>
        string Content();
    }
}
