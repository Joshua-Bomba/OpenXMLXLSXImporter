using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.ExcelGrid.Builders;
using OpenXMLXLXSImporter.ExcelGrid.Indexers;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid
{
    /// <summary>
    /// this will handle accessing the grid and iterateing over the spreadsheetgrid
    /// interate through all the cells for an entire column
    /// interate through all the cells for an entire row
    /// interate through all the cells
    /// </summary>
    public class SpreadSheetGrid : ISpreadSheetIndexersLock, IDisposable
    {
        private ISpreadSheetFileLockable _fileAccess;
        private ISheetProperties _sheetProperties;

        private Sheet _sheet;
        private WorksheetPart _workbookPart;
        private Worksheet _worksheet;
        private SheetData _sheetData;

        private Task _loadSpreadSheetData;

        private RowIndexer _rows;
        private ColumnIndexer _columns;
        private AsyncLock _accessorLock = new AsyncLock();
        private List<IIndexer> _indexers;

        private BlockingCollection<ICellProcessingTask> _cellTasks;

        AsyncLock ISpreadSheetIndexersLock.IndexerLock => _accessorLock;

        void ISpreadSheetIndexersLock.AddIndexer(IIndexer a)
        {
            _indexers.Add(a);
        }

        void ISpreadSheetIndexersLock.Spread(IIndexer a, ICellIndex b)
        {
            foreach(IIndexer i in _indexers)
            {
                if (i != a)
                    i.Spread(b);
            }
        }

        void ISpreadSheetIndexersLock.EnqueCell(ICellProcessingTask t)
        {
            _cellTasks.Add(t);
        }

        public SpreadSheetGrid(ISpreadSheetFileLockable fileAccess, ISheetProperties sheetProperties)
        {
            _fileAccess = fileAccess;
            _sheetProperties = sheetProperties;
            _indexers = new List<IIndexer>();
            _cellTasks = new BlockingCollection<ICellProcessingTask>();

            _loadSpreadSheetData = Task.Run(LoadSpreadSheetData);

            _rows = new RowIndexer(this);
            _columns = new ColumnIndexer(this);
        }

        protected async Task LoadSpreadSheetData()
        {
            await _fileAccess.ContextLock(async x =>
            {
                _sheet = x.GetSheet(_sheetProperties.Sheet);
                _workbookPart = x.WorkbookPart.GetPartById(_sheet.Id) as WorksheetPart;
                _worksheet = _workbookPart.Worksheet;
                _sheetData = _worksheet.Elements<SheetData>().First();
            });

            try
            {
                while (true)
                {
                    ICellProcessingTask t = _cellTasks.Take();
                }
            }
            catch (InvalidOperationException ex)
            {
                //queue is finished
            }

        }

        public void Dispose()
        {
            _loadSpreadSheetData.Wait();
        }

        public async Task ProcessInstruction(ISpreadSheetInstruction spreadSheetInstruction)
        {
            if (spreadSheetInstruction.ByColumn)
            {
                await _columns.ProcessInstruction(spreadSheetInstruction);
            }
            else
            {
                await _rows.ProcessInstruction(spreadSheetInstruction);
            }
        }

        //public async Task Add(ICellData cellData)
        //{
        //    using(await _lockRow.LockAsync())
        //    {
        //        if (!_rows.ContainsKey(cellData.CellRowIndex))
        //        {
        //            _rows[cellData.CellRowIndex] = new RowIndexer();
        //        }

        //        _rows[cellData.CellRowIndex].Add(cellData);
        //        if(_listeners != null)
        //        {
        //            //intresting so this will call each method and won't await till it's finished calling all of them
        //            //pretty handy neat pattern
        //            await Task.WhenAll(_listeners.Select(x => x.NotifyAsync(cellData)));
        //        }
        //    }
        //    using (await _lockColumn.LockAsync())
        //    {
        //        if (!_columns.ContainsKey(cellData.CellColumnIndex))
        //        {
        //            _columns[cellData.CellColumnIndex] = new ColumnIndexer(this);
        //        }

        //        _columns[cellData.CellColumnIndex].Add(cellData);
        //    }
        //}
    }
}
