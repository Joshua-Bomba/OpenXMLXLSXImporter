﻿using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter
{
    public interface IExcelImporter : IDisposable
    {
        Task<ISpreadSheetInstructionBuilder> GetSheetBuilder(string sheetName);
    }
}
