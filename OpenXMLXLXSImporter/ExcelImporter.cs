using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.FileAccess;
using OpenXMLXLSXImporter.Processing;

namespace OpenXMLXLSXImporter
{

    /// <summary>
    /// This class is for managing the import pass in a Stream to your excel file
    /// Implemented the ISheetProperties interface for each sheet you want to import
    /// add the ISheetProperties through the Add Method
    /// This will make it load the sheet in
    /// </summary>
    public class ExcelImporter : IExcelImporter
    {
        private XlsxDocumentFile _streamSheetFile;
        private SpreadSheetDequeManager dequeManager;
        public static IExcelImporter CreateExcelImporter(Stream stream) => new ExcelImporter(stream);
        public ExcelImporter(Stream stream)
        {
            _streamSheetFile = new XlsxDocumentFile(stream);
        }

        public Task Process(ISheetProperties sheet)
        {
            return Task.Run(async() =>
            {
                try
                {
                    SpreadSheetInstructionBuilder ssib = new SpreadSheetInstructionBuilder();
                    Task<IXlsxSheetFilePromise> gt = _streamSheetFile.LoadSpreadSheetData(sheet);
                    sheet.LoadConfig(ssib);
                    IXlsxSheetFilePromise g = await gt;
                    dequeManager = new SpreadSheetDequeManager();
                    SpreadSheetInstructionManager ssim = new SpreadSheetInstructionManager(dequeManager);
                    dequeManager.StartRequestProcessor(g);
                    await ssib.ProcessInstructions(ssim);
                    Task r = ssib.ProcessResults();
                    await sheet.ResultsProcessed(ssib);
                    await r;
                }
                catch(Exception ex)
                {
                    ExceptionDispatchInfo.Capture(ex).Throw();
                }

            });

        }

        public void Dispose()
        {
            dequeManager?.Finish();//we have all of our results processed we are finished adding
        }
    }
}
