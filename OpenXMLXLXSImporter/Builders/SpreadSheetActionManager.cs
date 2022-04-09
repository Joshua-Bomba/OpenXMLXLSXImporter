using OpenXMLXLSXImporter.CellData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Builders
{
    public class SpreadSheetActionManager : ISpreadSheetInstructionKey
    {
        //hmm I could use an event handler here but I think there stupid so I will not
        private IList<ISpreadSheetInstructionKey.SpreadSheetAction> _events;
        public SpreadSheetActionManager()
        {
            _events = null;
        }

        public void AddAction(ISpreadSheetInstructionKey.SpreadSheetAction evt)
        {
            if(_events == null)
                _events = new List<ISpreadSheetInstructionKey.SpreadSheetAction>();
            _events.Add(evt);
        }

        public async Task TriggerEvent(ISpreadSheetInstruction instruction)
        {
            if(_events != null)
            {
                IEnumerator<ISpreadSheetInstructionKey.SpreadSheetAction> ac = _events.GetEnumerator();

                bool exists = ac.MoveNext();
                if(exists)
                {
                    ISpreadSheetInstructionKey.SpreadSheetAction el = ac.Current;

                    IAsyncEnumerable<ICellData> cd = instruction.GetResults();
                    exists = ac.MoveNext();
                    el?.Invoke(cd);
                    if(exists)
                    {
                        do
                        {
                            el = ac.Current;
                            el?.Invoke(cd);
                        } while ((exists = ac.MoveNext()));
                    }
                }
            }
        }
    }
}
