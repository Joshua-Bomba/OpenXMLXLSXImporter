using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public interface ISpreadSheetInstruction
    {
        void AttachSpreadSheetInstructionManager(ISpreadSheetInstructionManager spreadSheetInstructionManager);

        void EnqueCell(IDataStoreLocked indexer);

        IAsyncEnumerable<ICellData> GetResults();
    }
}
