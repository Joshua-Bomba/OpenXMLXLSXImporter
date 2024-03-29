﻿using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
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
    public class FutureCell : IFutureCell, ICellProcessingTask, ICellIndex
    {

        private  ICellData _result;
        private IFutureUpdate _updater;
        private AsyncManualResetEvent _mre;
        private Exception _fail;
        private Task _enqued;
        public FutureCell(uint cellRowIndex, string cellColumnIndex, IFutureUpdate updater, IQueueAccess queueAccess)
        {
            _fail = null;
            CellRowIndex = cellRowIndex;
            CellColumnIndex = cellColumnIndex;
            _mre = new AsyncManualResetEvent(false);
            _updater = updater;
            _enqued = queueAccess.QueueCellProcessingTask(this);
        }
        public string CellColumnIndex { get; set; }

        public uint CellRowIndex { get; set; }
        public bool Processed => _mre.IsSet;

        public void Failure(Exception e)
        {
            _fail = e;
            _mre.Set();
        }

        public async Task<ICellData> GetData()
        {
            await _enqued;
            await _mre.WaitAsync();
            if(_fail != null )
            {
                throw _fail;
            }
            return _result;
        }

        public async Task Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
        {
            if(!Processed)
            {
                _result = await file.ProcessedCell(cellElement, index);
                _updater?.Update(_result);
                _mre.Set();
            }
        }
    }
}
