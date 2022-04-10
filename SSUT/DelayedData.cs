using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSUT
{
    public class DelayedData : MemoryStream
    {
        public DelayedData() : base()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            //System.Threading.Thread.Sleep(10);
            return base.Read(buffer, offset, count);
        }
    }
}
