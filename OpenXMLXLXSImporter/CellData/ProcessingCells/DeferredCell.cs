using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.FileAccess;
using OpenXMLXLSXImporter.Indexers;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public class DeferredCell : IFutureCell, ICellIndex
    {
        private Cell _deferredCell;
        private IIndexer _indexer;
        public DeferredCell(uint cellRowIndex, string cellColumnIndex, Cell cell)
        {
            CellRowIndex = cellRowIndex;
            CellColumnIndex = cellColumnIndex;
            _deferredCell = cell;
        }

        private class DeferredCellExecution : ISpreadSheetInstruction, ICellProcessingTask
        {
            private DeferredCell _deferredCell;
            private ICellData _result;
            private AsyncManualResetEvent _mre;
            public DeferredCellExecution(DeferredCell deferredCell)
            {
                _mre = new AsyncManualResetEvent(false);
                _deferredCell = deferredCell;
            }
            public bool IndexedByRow => true;

            public async Task EnqueCell(IIndexer indexer)
            {
                await indexer.Add(this);
            }

            public async IAsyncEnumerable<ICellData> GetResults()
            {
                await _mre.WaitAsync();
                yield return _result;
            }
            public void Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
            {
                _result = file.ProcessedCell(_deferredCell._deferredCell, _deferredCell);
                _mre.Set();
            }
        }

        public SpreadSheetInstructionManager InstructionManager { get; set; }

        public string CellColumnIndex { get; set; }
        public uint CellRowIndex { get; set; }

        public async Task<ICellData> GetData()
        {
            ISpreadSheetInstruction ins = new DeferredCellExecution(this);
            await InstructionManager.ProcessInstruction(ins);
            IAsyncEnumerator<ICellData> d= ins.GetResults().GetAsyncEnumerator();
            await d.MoveNextAsync();
            await _indexer.SetCell(d.Current);
            return d.Current;
        }

        public void SetIndexer(IIndexer indexer)
        {
            _indexer = indexer;
        }
    }
}
