﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.CellData.CellParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.FileAccess
{
    public class XlsxSheetFile : IXlsxSheetFile
    {
        private IXlsxDocumentFile _fileAccess;

        private string _sheetName;

        private IAsyncEnumerator<Row> _rowEnumerator;
        private Row _row;
        private bool _rowsLoadedIn ;

        private UInt32Value _rowIndexNullable;
        private uint _rowIndex;

        private IDictionary<uint, IEnumerator<Cell>> _rows;
        private ICellParser _parser;

        public XlsxSheetFile(IXlsxDocumentFile fileAccess, string sheetName, ICellParser parser)
        {
            _fileAccess = fileAccess;
            _sheetName = sheetName;
            _rowsLoadedIn = false;
            _parser = parser;
            _parser.AttachFileAccess(fileAccess);
            _rows = new Dictionary<uint, IEnumerator<Cell>>();
            _rowEnumerator = _fileAccess.GetRows(_sheetName).GetAsyncEnumerator();

        }

        async Task<IEnumerator<Cell>> IXlsxSheetFile.GetRow(uint desiredRowIndex)
        {
            //this will only load in up to the row we need
            //if we try loading in a row that does not exist then we will load all of them in
            //I'm assuming that data from here might still be in the file and we don't want to read things we don't need
            IEnumerator<Cell> cellEnumerator = null;
            if (_rows.ContainsKey(desiredRowIndex))
            {
                cellEnumerator = _rows[desiredRowIndex];
                return cellEnumerator;
            }
            IEnumerable<Cell> cellEnum;
            while (!_rowsLoadedIn)
            {
                if (await _rowEnumerator.MoveNextAsync())
                {
                    _row = _rowEnumerator.Current;
                    _rowIndexNullable = _row.RowIndex;
                    if (_rowIndexNullable.HasValue)
                    {
                        _rowIndex = _rowIndexNullable.Value;
                        cellEnum = _row.Elements<Cell>();
                        cellEnumerator = cellEnum.GetEnumerator();
                        _rows.Add(_rowIndex, cellEnumerator);
                        if (_rowIndex == desiredRowIndex)
                            return cellEnumerator;
                    }
                }
                else
                {
                    _rowsLoadedIn = true;
                    break;
                }
            }

            cellEnumerator = null;
            return cellEnumerator;
        }

        public async Task<uint> GetAllRows()
        {
            IEnumerator<Cell> cellEnumerator;
            IEnumerable<Cell> cellEnum;
            while (!_rowsLoadedIn)
            {
                if (await _rowEnumerator.MoveNextAsync())
                {
                    _row = _rowEnumerator.Current;
                    _rowIndexNullable = _row.RowIndex;
                    if (_rowIndexNullable.HasValue)
                    {
                        _rowIndex = _rowIndexNullable.Value;
                        cellEnum = _row.Elements<Cell>();
                        cellEnumerator = cellEnum.GetEnumerator();
                        _rows.Add(_rowIndex, cellEnumerator);
                    }
                }
                else
                {
                    _rowsLoadedIn = true;
                    break;
                }
            }
            return _rowIndex;
        }

        ICellData IXlsxSheetFile.ProcessedCell(Cell cellElement, ICellIndex index) => _parser.ProcessCell(cellElement, index);

        public static string GetColumnIndexByColumnReference(StringValue columnReference)
        {
            string v = columnReference.Value;
            for (int i = v.Length - 1; i >= 0; i--)
            {
                char c = v[i];
                if (!Char.IsNumber(c))
                {
                    return v.Substring(0, i + 1);
                }

            }
            return v;
        }
    }
}
