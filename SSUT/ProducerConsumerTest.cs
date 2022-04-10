using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nito.Collections;
using Nito.AsyncEx;

namespace SSUT
{
    [TestFixture]
    public class ProducerConsumerTest
    {
        [Test]
        public void TestingAsyncProducerConsumerQueue()
        {
            AsyncProducerConsumerQueue< string> apcq = new AsyncProducerConsumerQueue<string>();

            int records = 0;
            Task consumer =Task.Run(async () =>
            {
                try
                {
                    while (true)
                    {
                        await apcq.DequeueAsync();
                        records++;
                    }
                }
                catch (Exception ex)
                {

                }
            });

            //stuff that can't stop execution
            Task t1 = Task.Run(() =>
            {
                for(int i =0;i < 1000;i++)
                {
                    apcq.EnqueueAsync("test1");
                }
            });

            //stuff that can't stop execution
            Task t2 = Task.Run(() =>
            {
                for (int i = 0; i < 1000; i++)
                {
                    apcq.EnqueueAsync("test2");
                }
            });


            t1.Wait();
            t2.Wait();

            apcq.CompleteAdding();


            consumer.Wait();
        }
    }
}
