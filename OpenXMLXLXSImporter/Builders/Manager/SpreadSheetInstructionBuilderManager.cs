using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders.Managers
{

    public class SpreadSheetInstructionBuilderManager : ISpreadSheetInstructionBuilderManager, ISpreadSheetInstructionBuilderManagerInstructionBuilder
    {
        private string _sheet;
        private List<ISpreadSheetInstructionKey> _keys;
        private ISpreadSheetInstructionBuilder _builder;

        private Action<ISpreadSheetInstructionBuilderManagerInstructionBuilder> _execute;

        private IEnumerable<IAsyncEnumerable<ICellData>> _result;
        public SpreadSheetInstructionBuilderManager(string sheet,Action<ISpreadSheetInstructionBuilderManagerInstructionBuilder> ac)
        {
  
            _sheet = sheet;
            _keys = new List<ISpreadSheetInstructionKey>();
            _builder = null;
            _execute = ac;
            
        }
        string ISpreadSheetInstructionBuilderManager.Sheet => _sheet;

        public IEnumerable<IAsyncEnumerable<ICellData>> GetResults() => _result;

        ISpreadSheetInstructionBuilderManagerInstructionBuilder ISpreadSheetInstructionBuilderManagerInstructionBuilder.LoadSingleCell(uint row, string cell)
        {
            _keys.Add(_builder.LoadSingleCell(row, cell));
            return this;
        }

        void ISpreadSheetInstructionBuilderManager.LoadConfig(ISpreadSheetInstructionBuilder builder)
        {
            _builder = builder;
            _execute(this);
        }

        async Task ISpreadSheetInstructionBuilderManager.ResultsProcessed(ISpreadSheetQueryResults query)
        {
            _result = _keys.Select(x=>query.GetProcessedResults(x)).ToList();
        }
    }
}
