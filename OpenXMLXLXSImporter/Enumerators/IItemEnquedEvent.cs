using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid
{
    public interface IItemEnquedEvent
    {
        Task NotifyAsync(ICellData c);

        void FinishedLoading();
    }
}
