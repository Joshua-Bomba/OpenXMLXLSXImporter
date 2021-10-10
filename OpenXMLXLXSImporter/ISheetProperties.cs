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
        /// <summary>
        /// This is the Sheet name
        /// </summary>
        string Sheet { get; }

        /// <summary>
        /// this will be called when the SpreadSheetGrid is constructed for this sheet
        /// stored this in a variable and access it to access the spreadsheet
        /// </summary>
        /// <param name="grid"></param>
        void SetCellGrid(SpreadSheetGrid grid);
    }
}
