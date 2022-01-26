using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.Enumerators;
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
    public class SpreadSheetGrid
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


        public SpreadSheetGrid(ISpreadSheetFileLockable fileAccess, ISheetProperties sheetProperties)
        {
            _fileAccess = fileAccess;
            _sheetProperties = sheetProperties;


            _loadSpreadSheetData = Task.Run(LoadSpreadSheetData);

            _rows = new RowIndexer();
            _columns = new ColumnIndexer();
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
        }
        /// <summary>
        /// this will return a paticular row and cell if it's not avaliable it will wait till it's loaded or the timeout is reached
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="cellIndex"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public async Task<ICellData> FetchCell(uint rowIndex, string cellIndex)
        {
            ICellData cell = null;
            if(!await _rows.HasCell(rowIndex, cellIndex))//lock & unlock rows
            {
                await AddMissingCell(rowIndex, cellIndex);//_rows is unlocked here
            }
            return await _rows.GetCell(rowIndex, cellIndex);//lock & unlock rows
        }

        public async Task AddMissingCell(uint rowIndex, string cellIndex)
        {
            await _fileAccess.ContextLock(async x =>//File access is locked
            {
                //we are going to check if the cell has been added since we got the lock
                if (!await _rows.HasCell(rowIndex, cellIndex))//rows is locked
                {

                }
            });
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
