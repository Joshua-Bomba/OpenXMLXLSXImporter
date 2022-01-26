using Nito.AsyncEx;
using OpenXMLXLXSImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter.ExcelGrid.Indexers
{
    /// <summary>
    /// this manages access by row then cell
    /// </summary>
    public class RowIndexer : BaseIndexer
    {
        protected Dictionary<uint,Dictionary<string, ICellData>> _cells;

         
        public RowIndexer() : base()
        {
            _cells = new Dictionary<uint, Dictionary<string, ICellData>>();
        }
        /// <summary>
        /// this will add a cell called by the ExcelImporter
        /// </summary>
        /// <param name="cellData"></param>
        public async override Task Add(ICellData cellData)
        {

            //cellData.CellColumnIndex, cellData
            //_cells.Add();
        }

        public  async override Task<bool> HasCell(uint rowIndex, string cellIndex)
        {
            throw new NotImplementedException();
        }

        public async override Task<ICellData> GetCell(uint rowIndex, string cellIndex)
        {
            throw new NotImplementedException();
        }

    }
}
