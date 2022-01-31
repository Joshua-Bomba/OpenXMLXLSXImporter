using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.ExcelGrid.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid
{
    public class SpreadSheetLoadQueueManager
    {
        //new BlockingCollection<ICellProcessingTask>();
        public void Enque(ICellProcessingTask cell)
        {

        }

        public ICellProcessingTask Take()
        {
            throw new NotImplementedException();
        }
    }
}
