using Nito.AsyncEx;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Processing
{
    public interface ISpreadSheetIndexersLock
    {
        AsyncLock IndexerLock { get; }
        void AddIndexer(IIndexer a);
        void Spread(IIndexer a, ICellIndex b);
        Task EnqueCell(ICellProcessingTask t);   
    }
}
