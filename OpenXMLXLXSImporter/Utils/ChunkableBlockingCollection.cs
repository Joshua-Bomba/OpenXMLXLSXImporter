using Nito.AsyncEx;
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

        void ProcessQueue(ref Queue<T> items);

        void PostProcessing();

        void PreProcessing();
    }

    public interface IChunkableBlockingCollection<T>
    {
        AsyncLock Mutex { get; }
        void Enque(T item);

    }

    public class ChunkableBlockingCollection<T> : IChunkableBlockingCollection<T>
    {
        private AsyncLock _mutext;
        private BlockingCollection<T> _queue;
        private Queue<T> _chunkedItems;
        private IChunckBlock<T> _chunkBlock;
        public ChunkableBlockingCollection(IChunckBlock<T> chunkBlock)
        {
            _chunkBlock = chunkBlock;
            _mutext = new AsyncLock();//we will make the mre's inital state as true
            _queue = new BlockingCollection<T>();
            _chunkedItems = null;
            chunkBlock.Init(this);
        }

        public AsyncLock Mutex => _mutext;

        public void Enque(T item) => _queue.Add(item);

        private async Task Chunk()
        {
            _chunkBlock.PreProcessing();
            using(await _mutext.LockAsync())
            {
                BlockingCollection<T> dumpCollection = _queue;
                _queue = new BlockingCollection<T>();
                dumpCollection.CompleteAdding();
                if (_chunkedItems == null)
                {
                    _chunkedItems = new Queue<T>(dumpCollection);
                }
                else
                {
                    foreach (T item in dumpCollection)
                    {
                        _chunkedItems.Enqueue(item);
                    }
                }

                _chunkBlock.ProcessQueue(ref _chunkedItems);
            }
            _chunkBlock.PostProcessing();
            
        }

        public void Finish() => _queue.CompleteAdding();

        public async Task<T> Take()
        {
            if (_chunkedItems != null)
            {
                if (_chunkedItems.TryDequeue(out T item))
                {
                    if(_chunkBlock.ShouldPullAndChunk)
                    {
                        await Chunk();
                    }
                    return item;
                }
                else
                {
                    _chunkedItems = null;
                }
            }
            if (_chunkBlock.ShouldPullAndChunk)
            {
                await Chunk();
                if(_chunkedItems.TryDequeue(out T item))
                {
                    return item;
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
