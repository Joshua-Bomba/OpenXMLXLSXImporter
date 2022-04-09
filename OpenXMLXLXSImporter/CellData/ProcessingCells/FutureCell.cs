using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLSXImporter.FileAccess;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    internal class ResolveTracker<T>
    {
        public HashSet<T> set;
        private object l;

        public ResolveTracker()
        {
            set = new HashSet<T>();
            l = new object();
        }
        public void LogEnque(T c)
        {
            Task.Run(() =>
            {
                lock (l)
                {
                    set.Add(c);
                }
            });
        }

        public void LogDeque(T c)
        {
            Task.Run(() =>
            {
                lock (l)
                {
                    if(set.Contains(c))
                    {
                        set.Remove(c);
                    }
                }
            });
        }
    }


    public class FutureCell : IFutureCell, ICellProcessingTask, ICellIndex
    {
        internal static ResolveTracker<ICellProcessingTask> Tracker = new ResolveTracker<ICellProcessingTask>();

        private  ICellData _result;
        private IFutureUpdate _updater;
        private AsyncManualResetEvent _mre;
        public FutureCell(uint cellRowIndex, string cellColumnIndex, IFutureUpdate updater)
        {
            CellRowIndex = cellRowIndex;
            CellColumnIndex = cellColumnIndex;
            _mre = new AsyncManualResetEvent(false);
            _updater = updater;
            Tracker.LogEnque(this);
        }
        public string CellColumnIndex { get; set; }

        public uint CellRowIndex { get; set; }

        public async Task<ICellData> GetData()
        {
            await _mre.WaitAsync();
            return _result;
        }

        public void Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
        {
            try
            {
                Tracker.LogDeque(this);
                _result = file.ProcessedCell(cellElement, index);
                _updater?.Update(_result);
            }
            finally
            {
                _mre.Set();
            }

        }
    }
}
