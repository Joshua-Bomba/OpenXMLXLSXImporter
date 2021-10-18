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
    /// <summary>
    /// This class is for managing the import pass in a Stream to your excel file
    /// Implemented the ISheetProperties interface for each sheet you want to import
    /// add the ISheetProperties through the Add Method
    /// After your done adding them run the LoadSpreadSheetData method
    /// </summary>
    public class ExcelImporter : IDisposable
    {
        private SpreadsheetDocument _spreadsheet;
        private WorkbookPart _workbookPart;
        private SharedStringTable _sharedStringTable;
        private WorkbookStylesPart _workBookStyleParts;
        private Stylesheet _styleSheet;
        private CellFormats _cellFormats;
        private Stream _stream;
        private SpreadSheetGrid[] _grids;
        private List<ISheetProperties> _sheetProperties;
        private Task[] _processedWorkSheets;

        public ExcelImporter(Stream stream)
        {
            _stream = stream;
            _sheetProperties = new List<ISheetProperties>();
            _grids = null;
            _processedWorkSheets = null;
        }

        public void Add(ISheetProperties prop)
        {
            _sheetProperties.Add(prop);
        }
        /// <summary>
        /// This is for custom Cell Types like dates Cell with relations like text etc
        /// </summary>
        /// <param name="c"></param>
        /// <param name="cellData"></param>
        /// <returns></returns>
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

        private static string GetColumnIndexByColumnReference(StringValue columnReference)
        {
            string v = columnReference.Value;
            for(int i = v.Length - 1;i >= 0;i--)
            {
                char c = v[i];
                if (!Char.IsNumber(c))
                {
                    return v.Substring(0, i + 1);
                }

            }
            return v;
        }

        /// <summary>
        /// this will be called for each work sheet and will process each cell
        /// </summary>
        /// <param name="worksheetPart"></param>
        /// <param name="grid"></param>
        /// <returns></returns>
        protected async Task ProcessWorkSheet(WorksheetPart worksheetPart, SpreadSheetGrid grid)
        {
            await Task.Run(() => {
                List<Task> cellTask = new List<Task>();
                Worksheet ws = worksheetPart.Worksheet;
                SheetData sheetData = ws.Elements<SheetData>().First();
                IEnumerator<Row> enumerator = sheetData.Elements<Row>().GetEnumerator();
                while (enumerator.MoveNext())
                {
                    Row r = enumerator.Current;
                    IEnumerable<Cell> cells = r.Elements<Cell>();

                    foreach (Cell c in cells)
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
                                cellData.CellColumnIndex = GetColumnIndexByColumnReference(c.CellReference);
                                cellData.CellRowIndex = r.RowIndex.Value;
                                cellTask.Add(grid.Add(cellData));
                            }
                        }
                    }
                }
                cellTask.ForEach(x => x.Wait());
                grid.FinishedLoading();
            });

        }

        /// <summary>
        /// processes all the sheets passed in
        /// </summary>
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
            _processedWorkSheets = new Task[sheets.Length];
            _grids = new SpreadSheetGrid[sheets.Length];
            for (int i = 0; i < sheets.Length; i++)
            {
                WorksheetPart wp = _workbookPart.GetPartById(sheets[i].Id) as WorksheetPart;
                _grids[i] = new SpreadSheetGrid(_sheetProperties[i]);
                _processedWorkSheets[i] = this.ProcessWorkSheet(wp, _grids[i]);
            }
        }

        public void Dispose()
        {
            if(_processedWorkSheets != null)
            {
                foreach(Task t in _processedWorkSheets)
                    t?.Wait();
            }
            _spreadsheet.Dispose();
        }

        
    }
}
