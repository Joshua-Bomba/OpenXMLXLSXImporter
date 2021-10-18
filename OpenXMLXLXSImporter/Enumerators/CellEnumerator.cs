using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.ExcelGrid;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.Enumerators
{
    public class CellEnumerator : IAsyncEnumerator<ICellData>, IItemEnquedEvent
    {
        private ISpreadSheetGridCollectionAccessor _ssg;
        private ICellData _current;
        private AsyncCollection<ICellData> _data;
        internal CellEnumerator(ISpreadSheetGridCollectionAccessor ssg)
        {
            _ssg = ssg;
            _current = null;
            _data = null;
        }

        public ICellData Current => _current;

        public async ValueTask DisposeAsync()
        {

        }


        public async Task GetCurrentRows()
        {

            using (await _ssg.RowLock.LockAsync())
            {
                List<ICellData> d = new List<ICellData>();

                //initalize the AsyncCollection using a ConcurrentQueue which is inialized using the Linq statement which I think will produce
                //an Iterator which will iterate over each value to populate the ConcurrentQueue
                _data = new AsyncCollection<ICellData>(
                    new ConcurrentQueue<ICellData>(
                        _ssg.Rows
                        .SelectMany(x => x.Value.Contents)
                        .Select(x => x.Value)//pass in the IEnumerable of ICellData
                        )
                    );

                _ssg.EnqueListener(this);
            }
        }


        public async ValueTask<bool> MoveNextAsync()
        {
            try
            {
                if (_data == null)
                {
                    await GetCurrentRows();
                }

                ICellData c = await _data.TakeAsync();
                if (c != null)
                {
                    _current = c;
                    return true;
                }
            }
            catch
            {

            }
            return false;
        }

        public async Task NotifyAsync(ICellData c)
        {
            await _data.AddAsync(c);
        }
    }
}
