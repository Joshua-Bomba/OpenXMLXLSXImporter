using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.CellData
{
    /// <summary>
    /// a cell thats data is from a SharedStringTable
    /// </summary>
    public class CellDataRelation : BaseCellData
    {
        private string _content;
        private OpenXmlElement _sharedStringResult;
        public CellDataRelation(int index, SharedStringTable sst)
        {
            _content = null;
            Index = index;
            _sharedStringResult = sst.ElementAt(index);
            _content = _sharedStringResult.FirstChild.InnerText;
        }

        public int Index { get; set; }

        public override string Content() => _content;
    }
}
