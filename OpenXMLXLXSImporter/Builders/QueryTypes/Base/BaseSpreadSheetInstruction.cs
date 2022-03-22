using Nito.AsyncEx;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public abstract class BaseSpreadSheetInstruction : ISpreadSheetInstruction
    {
        protected AsyncManualResetEvent _mre;
        public BaseSpreadSheetInstruction() {
            _mre = new AsyncManualResetEvent(false);
        }

        public virtual bool ByColumn => false;

        async IAsyncEnumerable<ICellData> ISpreadSheetInstruction.GetResults()
        {
            await _mre.WaitAsync();
            IEnumerable<ICellIndex> cells = GetResults();
            foreach(ICellIndex cell in cells)
            {
                yield return await GetCellData(cell);
            }
          
        }

        public static async Task<ICellData> GetCellData(ICellIndex i)
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
