using Nito.AsyncEx;
using OpenXMLXLXSImporter.ExcelGrid;
using OpenXMLXLXSImporter.ExcelGrid.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.Enumerators
{
    internal interface ISpreadSheetGridCollectionAccessor : ICollectionAccessor
    {
        AsyncLock RowLock { get; }
        Dictionary<uint, RowIndexer> Rows { get; }
    }
}
