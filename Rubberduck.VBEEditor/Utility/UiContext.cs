using System;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.VBEditor.Utility
{
    public interface IUiContextProvider
    {
        bool IsExecutingInUiContext();
        SynchronizationContext UiContext { get; }
        TaskScheduler UiTaskScheduler { get; }
    }

    public class UiContextProvider : IUiContextProvider
    {
        private static SynchronizationContext Context { get; set; }
        private static TaskScheduler TaskScheduler { get; set; }
        private static readonly UiContextProvider UiContextInstance = new UiContextProvider();
        private static readonly object Lock = new object();

        private UiContextProvider() { }
        
        public static void Initialize()
        {
            lock (Lock)
            {
                if (Context != null)
                {
                    return;
                }

                Context = SynchronizationContext.Current;
                TaskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            }
        }

        public static UiContextProvider Instance() => UiContextInstance;

        public SynchronizationContext UiContext => Context;
        public TaskScheduler UiTaskScheduler => TaskScheduler;

        public bool IsExecutingInUiContext()
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
