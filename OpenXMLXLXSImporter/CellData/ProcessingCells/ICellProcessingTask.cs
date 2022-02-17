﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.CellData
{
    public interface ICellProcessingTask : IFutureCell
    {
        void Resolve(ICellData data);
    }
}
