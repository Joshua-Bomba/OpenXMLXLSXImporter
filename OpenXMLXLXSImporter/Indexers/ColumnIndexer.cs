using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public class ColumnIndexer: Dictionary<string, ICellIndex>
    { 
        public ColumnIndexer() :base()
        {

        }

        public LastColumn LastColumn { get; set; }
    }
}
