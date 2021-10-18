using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
{
    /// <summary>
    /// a common base class for the row & cell accessors
    /// </summary>
    public abstract class BaseSpreadSheetIndexer
    {
        //not much here yet idk how this will look like in the future
        protected SpreadSheetGrid _grid;
        public BaseSpreadSheetIndexer(SpreadSheetGrid grid)
        {
            _grid = grid;
        }
        public abstract void Add(ICellData celldata);

    }
}
