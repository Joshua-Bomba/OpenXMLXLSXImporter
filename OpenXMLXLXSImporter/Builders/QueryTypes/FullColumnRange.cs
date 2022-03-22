using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Indexers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public class FullColumnRange : ISpreadSheetInstruction
    {
        private uint _row;
        private string _startingColumn;
        private LastColumn _lastColumn;
        private Task<IEnumerable<ICellData>> _d;
        public FullColumnRange(uint row,string startingColumn = "A")
        {
            _row = row;
            _startingColumn = startingColumn;
        }

        public bool ByColumn => true;

        void ISpreadSheetInstruction.EnqueCell(IIndexer indexer)
        {
            _lastColumn = new LastColumn(_row);
            indexer.Add(_lastColumn);
            //_d = Task.Run(async () =>
            //{
            //    ICellData d = await _lastColumn.GetData();

                
            //});
            indexer.ProcessInstruction()
        }

        async IAsyncEnumerable<ICellData> ISpreadSheetInstruction.GetResults()
        {
            //ICellIndex ci = await d;

            ICellData[] cd = new ICellData[0];

            foreach(ICellIndex cell in cd)
            {
                yield return await BaseSpreadSheetInstruction.GetCellData(cell);
            }

        }
    }
}
