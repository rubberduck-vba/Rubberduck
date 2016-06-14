using System;
using System.Threading;
using System.Windows.Threading;

namespace Rubberduck.UI.Command.MenuItems
{
    public static class UiDispatcher
    {
        // thanks to Pellared on http://stackoverflow.com/a/12909070/1188513

        private static SynchronizationContext UiContext { get; set; }
        
        public static void Initialize()
        {
            if (UiContext == null)
            {
                UiContext = SynchronizationContext.Current;
            }
        }

        /// <summary>
        /// Invokes an action asynchronously on the UI thread.
        /// </summary>
        /// <param name="action">The action that must be executed.</param>
        public static void InvokeAsync(Action action)
        {
            CheckInitialization();

            UiContext.Post(x => action(), null);
        }

        /// <summary>
        /// Executes an action on the UI thread. If this method is called
        /// from the UI thread, the action is executed immendiately. If the
        /// method is called from another thread, the action will be enqueued
        /// on the UI thread's dispatcher and executed asynchronously.
        /// <para>For additional operations on the UI thread, you can get a
        /// reference to the UI thread's context thanks to the property
        /// <see cref="UiContext" /></para>.
        /// </summary>
        /// <param name="action">The action that will be executed on the UI
        /// thread.</param>
        public static void Invoke(Action action)
        {
            CheckInitialization();

            if (UiContext == SynchronizationContext.Current)
            {
                action();
            }
            else
            {
                InvokeAsync(action);
            }
        }

        private static void CheckInitialization()
        {
            if (UiContext == null) throw new InvalidOperationException("UiDispatcher is not initialized. Invoke Initialize() from UI thread first.");
        }

        public static void Shutdown()
        {
            Invoke(() => Dispatcher.CurrentDispatcher.InvokeShutdown());
        }
    }
}
