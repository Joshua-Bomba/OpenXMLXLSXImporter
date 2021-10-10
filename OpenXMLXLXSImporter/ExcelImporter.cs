using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLXSImporter.CellData;
using OpenXMLXLXSImporter.ExcelGrid;

namespace OpenXMLXLXSImporter
{
    public class ExcelImporter : IDisposable
    {
        private SpreadsheetDocument _spreadsheet;
        private WorkbookPart _workbookPart;
        private SharedStringTable _sharedStringTable;
        private WorkbookStylesPart _workBookStyleParts;
        private Stylesheet _styleSheet;
        private CellFormats _cellFormats;
        private Stream _stream;
        private SpreadSheetGrid[] _matrixes;
        private List<ISheetProperties> _sheetProperties;

        public ExcelImporter(Stream stream)
        {
            _stream = stream;
            _sheetProperties = new List<ISheetProperties>();
            _matrixes = null;
        }

        public void Add(ISheetProperties prop)
        {
            _sheetProperties.Add(prop);
        }

        protected bool ProcessCustomCell(Cell c, out ICellData cellData)
        {
            if (c.StyleIndex != null)
            {
                int index = int.Parse(c.StyleIndex.InnerText);
                CellFormat cellFormat = _cellFormats.ChildElements[index] as CellFormat;
                if (cellFormat != null)
                {
                    if (ExcelStaticData.DATE_FROMAT_DICTIONARY.ContainsKey(cellFormat.NumberFormatId))
                    {
                        if (!string.IsNullOrEmpty(c.CellValue.Text))
                        {
                            if (double.TryParse(c.CellValue.Text, out double cellDouble))
                            {

                                DateTime theDate = DateTime.FromOADate(cellDouble);
                                cellData = new CellDataDate { Date = theDate, DateFormat = ExcelStaticData.DATE_FROMAT_DICTIONARY[cellFormat.NumberFormatId] };
                                return true;
                            }
                        }
                    }
                }
            }

            if (c.DataType?.Value != null && c.DataType?.Value == CellValues.SharedString)
            {
                cellData = new CellDataRelation(int.Parse(c.CellValue.InnerText), _sharedStringTable);
                return true;
            }
            cellData = null;
            return false;
        }

        protected async Task ProcessWorkSheet(WorksheetPart worksheetPart, SpreadSheetGrid matrix)
        {
            await Task.Run(() => {
                Worksheet ws = worksheetPart.Worksheet;
                SheetData sheetData = ws.Elements<SheetData>().First();
                IEnumerator<Row> enumerator = sheetData.Elements<Row>().GetEnumerator();
                while (enumerator.MoveNext())
                {
                    Row r = enumerator.Current;
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        if (c?.CellValue != null)
                        {
                            ICellData cellData;
                            bool hasBeenProcessed = ProcessCustomCell(c, out cellData);
                            if (!hasBeenProcessed)
                            {
                                cellData = new CellDataContent { Text = c.CellValue.Text };
                            }

                            if (cellData != null)
                            {
                                cellData.CellColumnIndex = c.CellReference;
                                cellData.CellRowIndex = r.RowIndex.Value;
                                matrix.Add(cellData);
                            }
                        }
                    }
                }

                matrix.FinishedLoading();
            });

        }

        public void LoadSpreadSheetData()
        {
            string[] sheetNames = _sheetProperties.Select(x => x.Sheet).ToArray();
            Sheet[] sheets = new Sheet[sheetNames.Length];
            //_collection = new BlockingCollection<RowData>();
            _spreadsheet = SpreadsheetDocument.Open(_stream, false, new OpenSettings { AutoSave = false });
            _workbookPart = _spreadsheet.WorkbookPart;

            IEnumerator<Sheet> sheetEnumerator = _workbookPart.Workbook.Sheets.Cast<Sheet>().GetEnumerator();
            while (sheetEnumerator.MoveNext())
            {
                StringValue name = sheetEnumerator.Current.Name;
                if (name.HasValue)
                {
                    int index = Array.IndexOf(sheetNames, name.Value);
                    if (index >= 0)
                    {
                        sheets[index] = sheetEnumerator.Current;
                    }
                }
            }

            SharedStringTablePart sharedStringTablePart = _workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            _sharedStringTable = sharedStringTablePart.SharedStringTable;

            _workBookStyleParts = _workbookPart.WorkbookStylesPart;
            _styleSheet = _workBookStyleParts.Stylesheet;
            _cellFormats = _styleSheet.CellFormats;

            sheets = sheets.Where(x => x != null).ToArray();
            WorksheetPart[] worksheetsParts = new WorksheetPart[sheets.Length];
            Task[] processedWorkSheets = new Task[sheets.Length];
            _matrixes = new SpreadSheetGrid[sheets.Length];
            for (int i = 0; i < sheets.Length; i++)
            {
                WorksheetPart wp = _workbookPart.GetPartById(sheets[i].Id) as WorksheetPart;
                _matrixes[i] = new SpreadSheetGrid(_sheetProperties[i]);
                processedWorkSheets[i] = this.ProcessWorkSheet(wp, _matrixes[i]);
            }
        }

        public void Dispose()
        {
            _spreadsheet.Dispose();
        }

        
    }
}
