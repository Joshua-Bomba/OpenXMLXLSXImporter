using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Processing;
using System;
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

        private Dictionary<string, IXlsxSheetFilePromise> _loadedSheets;
        private Dictionary<string, Sheet> _sheetRef;

        public XlsxDocumentFile(Stream stream)
        {
            _stream = stream;
            _loadSpreadSheetData = LoadSpreadSheetDocuemntData();
            _loadedSheets = new Dictionary<string, IXlsxSheetFilePromise>();
        }

        //async Task<IXlsxDocumentFile> IXlsxDocumentFilePromise.GetLoadedFile()
        //{
        //    await _loadSpreadSheetData;//Ensure this is loaded first
        //    return this;
        //}

        async Task<Sheet> IXlsxDocumentFile.GetSheet(string sheetName)
        {
            await _loadSpreadSheetData;
            return _sheetRef[sheetName];
        } 

        async Task<CellFormat> IXlsxDocumentFile.GetCellFormat(int index)
        {
            await _loadSpreadSheetData;
            return _cellFormats.ChildElements[index] as CellFormat;
        } 

        async Task<WorksheetPart> IXlsxDocumentFile.GetWorkSheetPartById(string sheetID)
        {
            await _loadSpreadSheetData;
            return (_workbookPart.GetPartById(sheetID) as WorksheetPart);
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


       

        public async Task<IXlsxSheetFilePromise> LoadSpreadSheetData(string sheet)
        {
            await _loadSpreadSheetData;
            if (_sheetRef.ContainsKey(sheet))
            {
                if (!_loadedSheets.ContainsKey(sheet))
                {
                    //this is the first time we use this sheet
                    _loadedSheets[sheet] = new XlsxSheetFile(this, sheet);
                }
                return _loadedSheets[sheet];
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
