using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public interface ICellIndex
    {
        /// <summary>
        /// this will hold the CellColumnIndex Example A, F, AA, ect
        /// </summary>
        string CellColumnIndex { get; set; }
        /// <summary>
        /// this will store the row #
        /// </summary>
        uint CellRowIndex { get; set; }

        async ValueTask<ICellData> GetData()
        {
            if (this is IFutureCell fs)
            {
                return await fs.GetData();
            }
            else if (this is ICellData cd)
            {
                return cd;
            }
            else
            {
                return null;
            }
        }
    }
}
