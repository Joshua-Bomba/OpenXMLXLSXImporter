﻿using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public interface ISpreadSheetQueryResults
    {
        IAsyncEnumerable<ICellData> GetProcessedResults(ISpreadSheetInstructionKey key);
    }
}
