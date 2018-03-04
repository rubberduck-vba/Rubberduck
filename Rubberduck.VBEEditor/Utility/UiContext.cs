using System;
using System.Threading;

namespace Rubberduck.VBEditor.Utility
{
    public interface IUiContext
    {
        bool CheckContext();
    }

    public class UiContext : IUiContext
    {
        private static SynchronizationContext Context { get; set; }
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
                }
            }
        }

        public static UiContext Instance() => UiContextInstance;

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
