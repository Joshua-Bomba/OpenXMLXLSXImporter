﻿using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using OpenXMLXLSXImporter.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public class ColumnRange : BaseSpreadSheetInstruction
    {
        private uint _row;

        private uint _startColumnInt;
        private uint _endColumnInt;

        private ICellIndex[] _cellItems;


        public ColumnRange(uint row, string startColumn, string endColumn)
        {
            _row = row;
            _startColumnInt = ExcelColumnHelper.GetColumnStringAsIndex(startColumn);
            _endColumnInt = ExcelColumnHelper.GetColumnStringAsIndex(endColumn);
            ValidateRange();
            _cellItems = new ICellIndex[(_endColumnInt - _startColumnInt) + 1];
        }

        public ColumnRange(uint row, uint startColumn, uint endcolumn)
        {
            _row = row;
            _startColumnInt = startColumn;
            _endColumnInt = endcolumn;
            ValidateRange();
            _cellItems = new ICellIndex[(_endColumnInt - _startColumnInt) + 1];
        }

        private void ValidateRange()
        {

            if(_startColumnInt > _endColumnInt)
            {//🤷 Guess I could just flip it but meh
                throw new ArgumentException("The End Column has to be greater then the Start Column");
            }
        }

        protected override void EnqueCell(IDataStoreLocked indexer)
        {
            for(uint i = _startColumnInt;i <= _endColumnInt;i++)
            {
                string column = ExcelColumnHelper.GetColumnIndexAsString(i);

                _cellItems[i - _startColumnInt] = indexer.GetCell(_row, column);
            }
        }

        protected override IEnumerable<ICellIndex> GetResults()
        {
            foreach(ICellIndex el in _cellItems)
            {
                yield return el;
            }
        }
    }
}
