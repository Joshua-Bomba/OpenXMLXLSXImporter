using Nito.AsyncEx;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using OpenXMLXLSXImporter.Processing;
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
        protected ISpreadSheetInstructionManager _manager;
        public BaseSpreadSheetInstruction() {
            _mre = new AsyncManualResetEvent(false);
        }

        public void AttachSpreadSheetInstructionManager(ISpreadSheetInstructionManager spreadSheetInstructionManager)
        {
            _manager = spreadSheetInstructionManager;
        }

        async IAsyncEnumerable<ICellData> ISpreadSheetInstruction.GetResults()
        {
            await _mre.WaitAsync();
            IAsyncEnumerator<ICellIndex> cells = GetResults().GetAsyncEnumerator();
            ICellIndex i;
            while(await cells.MoveNextAsync())
            {
                i = cells.Current;
                if (i is IFutureCell fs)
                {
                    yield return await fs.GetData();
                }
                else if (i is ICellData cd)
                {
                    yield return cd;
                }
            }
        }


        protected abstract IAsyncEnumerable<ICellIndex> GetResults();

        async Task  ISpreadSheetInstruction.EnqueCell(IDataStore indexer)
        {
            await EnqueCell(indexer);
            _mre.Set();
        }


        protected abstract Task EnqueCell(IDataStore indexer);

    }
}
