using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.FileAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData.CellParsing
{
    public class DefaultCellParser : ICellParser
    {
        private IXlsxDocumentFile _fileAccess;
        public DefaultCellParser()
        {
            _fileAccess = null;
        }

        public void AttachFileAccess(IXlsxDocumentFile file)
        {
            _fileAccess = file;
        }

        public ICellData ProcessCell(Cell cellElement, ICellIndex index)
        {
            ICellData cellData;
            if (cellElement != null && cellElement.CellValue != null)
            {
                cellData = ProcessCell(cellElement);
            }
            else
            {
                cellData = new EmptyCell();
            }

            if (cellData != null && index != null)
            {
                //use the same value from the promise
                cellData.CellColumnIndex = index.CellColumnIndex;
                cellData.CellRowIndex = index.CellRowIndex;
            }

            return cellData;
        }

        /// <summary>
        /// This is for custom Cell Types like dates Cell with relations like text etc
        /// </summary>
        /// <param name="c"></param>
        /// <param name="cellData"></param>
        /// <returns></returns>
        private ICellData ProcessCell(Cell c)
        {
            string? backgroundColor = null;
            string? foregroundColor = null;
            if (c.StyleIndex != null)
            {

                int index = int.Parse(c.StyleIndex.InnerText);

                CellFormat cellFormat = _fileAccess.GetCellFormat(index).GetAwaiter().GetResult();
                Fill? fill = _fileAccess.GetCellFill(cellFormat).GetAwaiter().GetResult();
                PatternFill? patternFill = fill?.PatternFill;
                backgroundColor = patternFill?.ForegroundColor?.Rgb?.ToString();

                if (cellFormat != null)
                {

                    if (ExcelStaticData.DATE_FROMAT_DICTIONARY.ContainsKey(cellFormat.NumberFormatId))
                    {
                        if (!string.IsNullOrEmpty(c.CellValue.Text))
                        {
                            if (double.TryParse(c.CellValue.Text, out double cellDouble))
                            {
                                DateTime theDate = DateTime.FromOADate(cellDouble);
                                return new CellDataDate { Date = theDate, DateFormat = ExcelStaticData.DATE_FROMAT_DICTIONARY[cellFormat.NumberFormatId], BackgroundColor = backgroundColor, ForegroundColor = foregroundColor };
                            }
                        }
                    }
                }
            }

            if (c.DataType?.Value != null && c.DataType?.Value == CellValues.SharedString)
            {
                int index = int.Parse(c.CellValue.InnerText);
                OpenXmlElement sharedStringElement = _fileAccess.GetSharedStringTableElement(index).GetAwaiter().GetResult();
                return new CellDataRelation(index, sharedStringElement) { BackgroundColor = backgroundColor, ForegroundColor = foregroundColor };
            }
            return new CellDataContent { Text = c.CellValue.Text, BackgroundColor = backgroundColor, ForegroundColor = foregroundColor };
        }
    }
}
