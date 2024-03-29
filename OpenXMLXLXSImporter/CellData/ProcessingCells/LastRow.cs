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
    public class LastRow : ICellProcessingTask
    {
        private ICellIndex _result;
        private AsyncManualResetEvent _mre;
        private Exception _fail;
        private Task _enqued;
        public LastRow(IQueueAccess queueAccess)
        {
            _fail = null;
            _mre = new AsyncManualResetEvent(false);
            _enqued = queueAccess.QueueCellProcessingTask(this);
        }

        public bool Processed => _mre.IsSet;

        public void Failure(Exception e)
        {
            _fail = e;
            _mre.Set();
        }

        public async Task<ICellIndex> GetIndex()
        {
            await _enqued;
            await _mre.WaitAsync();
            if(_fail != null)
            {
                throw _fail;
            }
            return _result;
        }
        public async Task Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
        {
            if(!Processed)
            {
                _result = index;
                _mre.Set();
            }
        }
    }
}
