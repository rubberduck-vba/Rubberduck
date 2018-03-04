using System;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.VBEditor.Utility
{
    public interface IUiContext
    {
        bool CheckContext();
        TaskScheduler UiTaskScheduler { get; }
    }

    public class UiContext : IUiContext
    {
        private static SynchronizationContext Context { get; set; }
        private static TaskScheduler TaskScheduler { get; set; }
        private static readonly UiContext UiContextInstance = new UiContext();
        private static readonly object Lock = new object();

        private UiContext() { }
        
        public static void Initialize()
        {
            lock (Lock)
            {
                if (Context == null)
                {
                    Context = SynchronizationContext.Current;
                    TaskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
                }
            }
        }

        public static UiContext Instance() => UiContextInstance;

        public TaskScheduler UiTaskScheduler => TaskScheduler;

        public bool CheckContext()
        {
            lock (Lock)
            {
                if (Context == null)
                {
                    throw new InvalidOperationException(
                        "UiContext is not initialized. Invoke Initialize() from UI thread first.");
                }

                return Context == SynchronizationContext.Current;
            }
        }
    }
}
