using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    public interface IDataStore
    {
        Task ProcessInstruction(ISpreadSheetInstruction instruction);

        Task<ICellIndex> GetCell(uint rowIndex, string cellIndex,Func<ICellProcessingTask> newCell = null);

        Task<ICellIndex> GetLastColumn(uint rowIndex);
        Task<ICellIndex> GetLastRow(uint columnIndex);

        Task SetCell(ICellIndex d);
        Task SetCells(IEnumerable<ICellIndex> cells);

        Task QueueNonIndexedCell(ICellProcessingTask t);

        
    }
}
