using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.FileAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public interface ICellProcessingTask
    {
        bool Processed { get; }
        void Resolve(ICellData data);
        void Resolve(IXlsxSheetFile file, Cell cellElement,ICellIndex index);
        void Failure(Exception e);
    }
}
