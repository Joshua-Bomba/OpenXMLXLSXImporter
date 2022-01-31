using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.CellData
{
    public class FutureCell : IFutureCell, ICellProcessingTask
    {
        private ICellData finalcellResult;
        public FutureCell(uint cellRowIndex, string cellColumnIndex)
        {
            CellRowIndex = cellRowIndex;
            CellColumnIndex = cellColumnIndex;
        }
        public string CellColumnIndex { get; set; }

        public uint CellRowIndex { get; set; }

        public Task<ICellData> GetData()
        {
            throw new NotImplementedException();
        }
    }
}
