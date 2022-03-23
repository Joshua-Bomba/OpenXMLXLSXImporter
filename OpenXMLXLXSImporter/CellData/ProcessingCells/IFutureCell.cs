using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public  interface IFutureCell
    {
        Task<ICellData> GetData();
        void SetIndexer(IIndexer indexer);
    }
}
