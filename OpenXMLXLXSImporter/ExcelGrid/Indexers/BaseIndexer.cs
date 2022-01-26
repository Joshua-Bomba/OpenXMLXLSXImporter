using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
{
    public abstract class BaseIndexer : IIndexer
    {
        protected readonly AsyncLock _lock = new AsyncLock();
        public BaseIndexer()
        {

        }

        public abstract Task Add(ICellIndex cellData);

        public abstract Task<ICellData> GetCell(uint rowIndex, string cellIndex);

        public abstract Task<bool> HasCell(uint rowIndex, string cellIndex);
    }
}
