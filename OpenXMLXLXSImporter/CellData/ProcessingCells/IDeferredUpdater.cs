using OpenXMLXLSXImporter.FileAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public  interface IDeferredUpdater
    {
        Task<Dictionary<DeferredCell, ICellProcessingTask>> AddDeferredCells(IEnumerable<DeferredCell> dc);
    }
}
