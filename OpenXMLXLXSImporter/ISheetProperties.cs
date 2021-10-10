using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLXSImporter.ExcelGrid;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter
{
    public interface ISheetProperties
    {
        string Sheet { get; }

        void SetCellMatrix(SpreadSheetGrid matrix);
    }
}
