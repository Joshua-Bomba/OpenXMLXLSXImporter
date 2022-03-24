using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public interface IIndexer
    {
        Task ProcessInstruction(ISpreadSheetInstruction instruction);

        Task<ICellIndex> GetCell(uint rowIndex, string cellIndex,Func<ICellProcessingTask> newCell = null);

        Task SetCell(ICellData d);

        Task QueueNonIndexedCell(ICellProcessingTask t);

        void Spread(ICellIndex cell);
    }
}
