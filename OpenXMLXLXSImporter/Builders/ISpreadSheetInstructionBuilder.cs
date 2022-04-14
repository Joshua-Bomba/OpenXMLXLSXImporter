using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public  interface ISpreadSheetInstructionBuilder
    {
        ISpreadSheetInstructionRunner Runner { get; }

        ISpreadSheetInstructionBundler GetBundler();

        Task BundleRequeset(IEnumerable<ISpreadSheetInstruction> sheetInstructions);
    }
}
