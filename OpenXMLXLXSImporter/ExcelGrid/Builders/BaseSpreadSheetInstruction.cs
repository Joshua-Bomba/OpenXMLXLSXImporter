using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.ExcelGrid.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Builders
{
    public abstract class BaseSpreadSheetInstruction : ISpreadSheetInstruction
    {
        protected AsyncManualResetEvent _mre;
        public BaseSpreadSheetInstruction() {}

        public virtual bool ByColumn => false;

        async Task<IEnumerable<Task<ICellData>>> ISpreadSheetInstruction.GetResults()
        {
            await _mre.WaitAsync();
            IEnumerable<ICellIndex> cells = GetResults();
            return cells.Select(GetCellData).ToArray();//iterate through the entire collection
          
        }

        private static async Task<ICellData> GetCellData(ICellIndex i)
        {
            if (i is IFutureCell fs)
            {
                return await fs.GetData();
            }
            else if(i is ICellData cd)
            {
                return cd;
            }
            throw new InvalidOperationException();
        }


        protected abstract IEnumerable<ICellIndex> GetResults();

        void  ISpreadSheetInstruction.EnqueCell(IIndexer indexer)
        {
            _mre.Set();
            EnqueCell(indexer);
        }


        protected abstract void EnqueCell(IIndexer indexer);
    }
}
