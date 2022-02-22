using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Utils
{
    public interface IChunckBlock<T>
    {
        void Init(ChunkableBlockingCollection<T> collection);

        bool ShouldPullAndChunk { get; }

        bool KeepQueueLockedForDump();

        void QueueDumpped(ref List<T> items);

        void PostProcessing();

        void PreProcessing();
    }


    public class ChunkableBlockingCollection<T>
    {
        private ManualResetEventSlim _mre;
        private BlockingCollection<T> _queue;
        private IEnumerator<T> _chunkedItems;
        private IChunckBlock<T> _chunkBlock;
        public ChunkableBlockingCollection(IChunckBlock<T> chunkBlock)
        {
            _chunkBlock = chunkBlock;
            _mre = new ManualResetEventSlim(true);//we will make the mre's inital state as true
            _queue = new BlockingCollection<T>();
            _chunkedItems = null;
            chunkBlock.Init(this);
        }

        public void Enque(T item)
        {
            _mre.Wait();
            _queue.Add(item);
        }

        private void SetupChunk()
        {
            _chunkBlock.PreProcessing();
            _mre.Reset();
            BlockingCollection<T> dumpCollection = _queue;
            _queue = new BlockingCollection<T>();
            bool keepLocked = _chunkBlock.KeepQueueLockedForDump();
            if (!keepLocked)
            {
                _mre.Set();
            }
            dumpCollection.CompleteAdding();
            List<T> queueOutput = dumpCollection.ToList();
            _chunkBlock.QueueDumpped(ref queueOutput);
            if (keepLocked)
            {
                _mre.Set();
            }
            _chunkBlock.PostProcessing();
            _chunkedItems = queueOutput.GetEnumerator();
        }

        public void Finish() => _queue.CompleteAdding();

        public T Take()
        {
            if (_chunkedItems != null)
            {
                if (_chunkedItems.MoveNext())
                {
                    return _chunkedItems.Current;
                }
                else
                {
                    _chunkedItems = null;
                }
            }
            if (_chunkBlock.ShouldPullAndChunk)
            {
                SetupChunk();
                if (_chunkedItems.MoveNext())
                {
                    return _chunkedItems.Current;
                }
                else
                {
                    _chunkedItems = null;
                }
            }
            return _queue.Take();
        }
    }
}
