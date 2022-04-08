﻿using DocumentFormat.OpenXml.Spreadsheet;
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
    public class LastRow : ICellProcessingTask, IFutureCell
    {
        public string _column;
        private ICellData _result;
        private AsyncManualResetEvent _mre;
        public LastRow(string column)
        {
            _mre = new AsyncManualResetEvent(false);
            _column = column;
        }

        public async Task<ICellData> GetData()
        {
            await _mre.WaitAsync();
            return _result;
        }
        public void Resolve(IXlsxSheetFile file, Cell cellElement, ICellIndex index)
        {
            _result = file.ProcessedCell(cellElement,index);
            _mre.Set();
        }

        public void SetDataStore(IDataStore indexer)
        {

        }
    }
}
