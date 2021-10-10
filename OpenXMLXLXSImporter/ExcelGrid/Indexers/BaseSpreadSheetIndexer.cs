using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
{
    public abstract class BaseSpreadSheetIndexer
    {
        public BaseSpreadSheetIndexer()
        {

        }
        public abstract void Add(ICellData celldata);

    }
}
