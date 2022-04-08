using Nito.AsyncEx;
using OpenXMLXLSXImporter.CellData;
using OpenXMLXLSXImporter.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Indexers
{
    /// <summary>
    /// this manages access by row then cell
    /// </summary>
    public class RowIndexer : Dictionary<uint,ColumnIndexer>
    {
        public RowIndexer()
        {
            
        }

        public LastRow LastRow { get; set; }

        public  ICellIndex Get(uint rowIndex, string cellIndex)
        {
            if (base.ContainsKey(rowIndex) && base[rowIndex].ContainsKey(cellIndex))
            {
                return base[rowIndex][cellIndex];
            }
            return null;
        }

        public void Set(ICellIndex cellData)
        {
            if (!base.ContainsKey(cellData.CellRowIndex))
            {
                base.Add(cellData.CellRowIndex, new ColumnIndexer());
            }
            base[cellData.CellRowIndex][cellData.CellColumnIndex] = cellData;
        }

        public void Add(ICellIndex cellData)
        {
            if (!base.ContainsKey(cellData.CellRowIndex))
            {
                base.Add(cellData.CellRowIndex, new ColumnIndexer());
            }
            base[cellData.CellRowIndex].Add(cellData.CellColumnIndex, cellData);
        }
    }
}
