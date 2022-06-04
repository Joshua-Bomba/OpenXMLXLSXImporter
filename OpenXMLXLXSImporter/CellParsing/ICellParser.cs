using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellParsing
{
    public interface ICellParser
    {
        Task<ICellData> ProcessCell(Cell cellElement);
    }
}
