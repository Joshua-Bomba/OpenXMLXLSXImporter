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

        Task PreQueueProcessing();
        void ProcessQueue(ref Queue<T> items);
        Task PostQueueProcessing();

        Task PostLockProcessing();

        Task PreLockProcessing();
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

        public void Enque(T item)
        {
            if(!_queue.IsAddingCompleted)
            {
                _queue.Add(item);
            }
            else
            {
                throw new Exception("No More requests are being accepted");
            }

        } 

        private async Task Chunk()
        {
            await _chunkBlock.PreLockProcessing();
            using(await _mutext.LockAsync())
            {
                await _chunkBlock.PreQueueProcessing();
                BlockingCollection<T> dumpCollection = _queue;
                _queue = new BlockingCollection<T>();
                if(dumpCollection.IsAddingCompleted)
                {
                    _queue.CompleteAdding();
                }
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
                await _chunkBlock.PostQueueProcessing();
            }
            await _chunkBlock.PostLockProcessing();
            
        }

        public BlockingCollection<T> Finish()
        {
            using(Mutex.Lock())
            {
                _queue.CompleteAdding();
            }
            return _queue;
        }

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
