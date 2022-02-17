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
    public interface ISpreadSheetIndexersLock
    {
        AsyncLock IndexerLock { get; }
        void AddIndexer(IIndexer a);
        void Spread(IIndexer a, ICellIndex b);
        void EnqueCell(ICellProcessingTask t);   
    }
}
