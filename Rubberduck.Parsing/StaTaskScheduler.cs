using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.Parsing
{
    public sealed class StaTaskScheduler : TaskScheduler, IDisposable
    {
        private readonly List<Thread> _threads;
        private BlockingCollection<Task> _tasks;

        public StaTaskScheduler()
            : this(1) { }

        public StaTaskScheduler(int concurrencyLevel)
        {
            if (concurrencyLevel < 1)
            {
                throw new ArgumentOutOfRangeException("concurrencyLevel");
            }

            _tasks = new BlockingCollection<Task>();
            _threads = Enumerable.Range(0, concurrencyLevel).Select(i =>
            {
                var thread = new Thread(() =>
                {
                    foreach (var task in _tasks.GetConsumingEnumerable())
                    {
                        TryExecuteTask(task);
                    }
                });
                thread.IsBackground = true;
                thread.SetApartmentState(ApartmentState.STA);
                return thread;
            }).ToList();

            _threads.ForEach(thread => thread.Start());
        }

        protected override void QueueTask(Task task)
        {
            _tasks.Add(task);
        }

        protected override bool TryExecuteTaskInline(Task task, bool taskWasPreviouslyQueued)
        {
            // todo: figure out how to implement
            return false;
        }

        protected override IEnumerable<Task> GetScheduledTasks()
        {
            return _tasks.ToArray();
        }

        public override int MaximumConcurrencyLevel
        {
            get { return _threads.Count; }
        }

        public void Dispose()
        {
            if (_tasks != null)
            {
                _tasks.CompleteAdding();
                foreach (var thread in _threads)
                {
                    thread.Join();
                }

                _tasks.Dispose();
                _tasks = null;
            }
        }
    }
}
