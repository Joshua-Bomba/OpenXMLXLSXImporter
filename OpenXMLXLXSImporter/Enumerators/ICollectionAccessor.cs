using OpenXMLXLXSImporter.ExcelGrid;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.Enumerators
{
    internal interface ICollectionAccessor
    {
        void EnqueListener(IItemEnquedEvent listener);
    }
}
