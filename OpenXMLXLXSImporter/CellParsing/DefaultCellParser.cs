using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.FileAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellParsing
{
    public class DefaultCellParser : ICellParser
    {
        protected IXlsxDocumentFile _fileAccess;

        protected Cell cell;
        protected CellFormat cellFormat;
        protected Fill? fill;
        protected PatternFill? patternFill;

        protected BaseCellData resultCell;

        public DefaultCellParser(IXlsxDocumentFile file)
        {
            _fileAccess = file;
        }

        protected virtual async Task ProcessCellFormating()
        {
            if (cell.StyleIndex != null)
            {
                int index = int.Parse(cell.StyleIndex.InnerText);
                cellFormat = await _fileAccess.GetCellFormat(index);
            }
        }

        protected virtual async Task ProcessCellColor()
        {
            if(cellFormat != null)
            {
                fill = await _fileAccess.GetCellFill(cellFormat);
                if (fill.PatternFill.ForegroundColor != null)
                {
                    if (fill.PatternFill.ForegroundColor.Rgb != null)
                    {
                        string colour = fill.PatternFill.ForegroundColor.Rgb.Value;
                        if(colour.Length >= 6)
                        {
                            resultCell.BackgroundColor = colour[^6..];
                        }
                        else
                        {
                            resultCell.BackgroundColor = colour;
                        }
                    }
                    else if (fill.PatternFill.ForegroundColor.Theme != null)
                    {
                        var color = await _fileAccess.GetColorByThemeIndex(fill.PatternFill.ForegroundColor.Theme.Value);
                        if(color.RgbColorModelHex != null)
                        {
                            resultCell.BackgroundColor = color.RgbColorModelHex.Val;
                        }
                        else if (color.SystemColor != null)
                        {
                            resultCell.BackgroundColor = color.SystemColor.LastColor.Value;
                        }
                        else
                        {
                            throw new NotImplementedException();
                        }

                        
                    }
                }
            }
        }

        protected virtual void AttachReferenceCell()
        {
            if(cell.CellFormula != null)
            {
                resultCell.CellFormula = cell.CellFormula.Text;
            }
        }

        protected virtual void CheckAndProcessCellIfDataDate()
        {
            if (cellFormat != null)
            {
                if (ExcelStaticData.DATE_FROMAT_DICTIONARY.ContainsKey(cellFormat.NumberFormatId))
                {
                    if (!string.IsNullOrEmpty(cell.CellValue.Text))
                    {
                        if (double.TryParse(cell.CellValue.Text, out double cellDouble))
                        {
                            DateTime theDate = DateTime.FromOADate(cellDouble);
                            resultCell = new CellDataDate { Date = theDate, DateFormat = ExcelStaticData.DATE_FROMAT_DICTIONARY[cellFormat.NumberFormatId] };
                        }
                    }
                }
            }
        }

        protected virtual async Task CheckAndProcessCellIfSharedString()
        {
            if (cell.DataType?.Value != null && cell.DataType?.Value == CellValues.SharedString)
            {
                int index = int.Parse(cell.CellValue.InnerText);
                OpenXmlElement sharedStringElement = await _fileAccess.GetSharedStringTableElement(index);
                resultCell =  new CellDataRelation(index, sharedStringElement);
            }
        }

        protected virtual void CreateGeneralContentCell()
        {
            resultCell = new CellDataContent { Text = cell.CellValue.Text  };
        }

        public async Task<ICellData> ProcessCell(Cell c)
        {
            cell = c;
            await ProcessCellFormating();
            CheckAndProcessCellIfDataDate();
            if(resultCell == null)
            {
                await CheckAndProcessCellIfSharedString();
                if (resultCell == null)
                {
                    CreateGeneralContentCell();
                }
            }
            await ProcessCellColor();
            AttachReferenceCell();

            return resultCell;
        }

    }
}
