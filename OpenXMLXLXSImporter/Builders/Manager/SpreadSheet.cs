using Nito.AsyncEx;
using OpenXMLXLSXImporter.Builders;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.FileAccess;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders.Managers
{

    public class SpreadSheet :  ISpreadSheet
    {
        private IXlsxSheetFile _file;
        private SpreadSheetDequeManager dequeManager;
        private SpreadSheetInstructionManager ssim;

        public SpreadSheet(IXlsxSheetFile file)
        {
            _file = file;
            
        }

        public void Spread(ICellIndex c)
        {
            ISpreadSheetIndexersLock whatever = ssim;
            whatever.Spread(null, c);
        }

        public void Dispose()
        {
            dequeManager?.Finish();
        }

        //string ISpreadSheet.Sheet => _sheetName;

        //public async IAsyncEnumerable<IEnumerable<Task<ICellData>>> GetResults()
        //{
        //    foreach(Task<IEnumerable<Task<ICellData>>> k in _result)
        //        yield return await k;
        //}

        async Task<ICellData> ISpreadSheet.LoadSingleCell(uint row, string cell)
        {
            //_keys.Add(_builder.LoadSingleCell(row, cell));
            //return this;
            throw new NotImplementedException();
        }

        //void ISpreadSheet.LoadConfig(ISpreadSheetInstructionBuilder builder)
        //{
        //    _builder = builder;
        //    _execute(this);
        //}

        //async Task ISpreadSheet.ResultsProcessed(ISpreadSheetQueryResults query)
        //{
        //    _result = _keys.Select(x=>query.GetResults(x)).ToList();
        //}
    }
}
