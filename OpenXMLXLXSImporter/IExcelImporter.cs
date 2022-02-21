using OpenXMLXLSXImporter.Builders.Managers;
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
        Task Process(ISpreadSheetInstructionBuilderManager prop);

        IAsyncEnumerable<IEnumerable<Task<ICellData>>> ProcessAndGetAsyncCollection(string sheetName, Action<ISpreadSheetInstructionBuilderManagerInstructionBuilder> builder);

        Task<List<List<ICellData>>> ProcessAndGetListAsync(string sheetName, Action<ISpreadSheetInstructionBuilderManagerInstructionBuilder> builder);

    }
}
