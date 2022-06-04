using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.CellParsing;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.FileAccess
{
    public class XlsxDocumentFile : IXlsxDocumentFile, IDisposable
    {
        private Stream _stream;
        //The Sheet
        private SpreadsheetDocument _spreadsheet;
        //WorkBookParts
        private WorkbookPart _workbookPart;
        private SharedStringTablePart _sharedStringTablePart;
        private SharedStringTable _sharedStringTable;
        private WorkbookStylesPart _workBookStyleParts;
        private Stylesheet _styleSheet;
        private CellFormats _cellFormats;
        //The Task that Loads in the SpreadSheetDocumentData
        private Task _loadSpreadSheetData;

        private ConcurrentDictionary<string, IXlsxSheetFile> _loadedSheets;
        private Dictionary<string, Sheet> _sheetRef;

        public XlsxDocumentFile(Stream stream)
        {
            _stream = stream;
            _loadSpreadSheetData = LoadSpreadSheetDocuemntData();
            _loadedSheets = new ConcurrentDictionary<string, IXlsxSheetFile>();
        }

        //async Task<IXlsxDocumentFile> IXlsxDocumentFilePromise.GetLoadedFile()
        //{
        //    await _loadSpreadSheetData;//Ensure this is loaded first
        //    return this;
        //}


        public async Task<Fill> GetCellFill(CellFormat cellFormat)
        {
            await _loadSpreadSheetData;
            if(cellFormat.FillId.HasValue)
            {
                return (Fill)_styleSheet.Fills.ChildElements[(int)cellFormat.FillId.Value];
            }
            return null;
        }


        async Task<CellFormat> IXlsxDocumentFile.GetCellFormat(int index)
        {
            await _loadSpreadSheetData;
            return _cellFormats.ChildElements[index] as CellFormat;
        } 

        async IAsyncEnumerable<Row> IXlsxDocumentFile.GetRows(string sheetName)
        {
            await _loadSpreadSheetData;
            Sheet sheet = _sheetRef[sheetName];
            WorksheetPart part = (_workbookPart.GetPartById(sheet.Id) as WorksheetPart);
            SheetData d =  part.Worksheet.Elements<SheetData>().First();
            IEnumerator<Row> r = d.Elements<Row>().GetEnumerator();
            while (r.MoveNext())
            {
                yield return r.Current;
            }
        }
        async Task<OpenXmlElement> IXlsxDocumentFile.GetSharedStringTableElement(int index)
        {
            await _loadSpreadSheetData;
            return _sharedStringTable.ElementAt(index);
        } 

        /// <summary>
        /// Common Reusable Parts of the WorkSheet
        /// </summary>
        /// <returns></returns>
        private Task LoadSpreadSheetDocuemntData()
        {
            return Task.Run(() =>
            {
                //Replace this stuff with OpenXmlReader so it does not load everything in at once otherwise there is pretty much no point in how I designed this.
                //...Not now maybe  later... My SpreadSheet is big but not super massive so it's fine for now
                _spreadsheet = SpreadsheetDocument.Open(_stream, false, new OpenSettings { AutoSave = false });
                _workbookPart = _spreadsheet.WorkbookPart;

                _sharedStringTablePart = _workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                _sharedStringTable = _sharedStringTablePart.SharedStringTable;

                _workBookStyleParts = _workbookPart.WorkbookStylesPart;
                _styleSheet = _workBookStyleParts.Stylesheet;
                _cellFormats = _styleSheet.CellFormats;

                //this will get all the sheetNames
                _sheetRef = _workbookPart.Workbook.Sheets.Cast<Sheet>().Where(x => x.Name.HasValue).ToDictionary(x => x.Name.Value);
            });
        }


       

        public async Task<IXlsxSheetFile> LoadSpreadSheetData(string sheet, ICellParser parser)
        {
            await _loadSpreadSheetData;
            if (_sheetRef.ContainsKey(sheet))
            {
                return _loadedSheets.GetOrAdd(sheet, x=> new XlsxSheetFile(this, sheet, parser));
            }
            return null;
        }

        public void Dispose()
        {
            //if(_loadedSheets != null)
            //{
            //    foreach(Task t in _loadedSheets)
            //        t?.Wait();
            //}
            _loadSpreadSheetData.Wait();
            _spreadsheet.Dispose();
        }
    }
}
