using System;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.UIContext
{
    public static class UiDispatcher
    {
        // thanks to Pellared on http://stackoverflow.com/a/12909070/1188513

        private static SynchronizationContext UiContext { get; set; }
        private static TaskScheduler UiTaskScheduler { get; set; }

        public static void Initialize()
        {
            if (UiContext == null)
            {
                UiContext = SynchronizationContext.Current;
                UiTaskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
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
        /// thread</param>
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

        /// <summary>
        /// Starts a task on the ui thread.
        /// </summary>
        /// <param name="action">The action that will be executed on the UI
        /// thread.</param>
        /// <param name="token">Optional cancellation token</param>
        /// <param name="options">Optional TaskCreationOptions</param>
        /// <returns></returns>
        public static Task StartTask(Action action, CancellationToken token, TaskCreationOptions options = TaskCreationOptions.None)
        {
            CheckInitialization();

            return Task.Factory.StartNew(action, token, options, UiTaskScheduler);
        }

        //This separate overload is necessary because CancellationToken.None is not a compile-time constant and thus cannot be used as default value.
        public static Task StartTask(Action action, TaskCreationOptions options = TaskCreationOptions.None)
        {
            return StartTask(action, CancellationToken.None, options);
        }


        /// <summary>
        /// Starts a task returning a value on the ui thread.
        /// </summary>
        /// <param name="func">The function that will be executed on the UI
        /// thread.</param>
        /// <param name="token">Optional cancellation token</param>
        /// <param name="options">Oprional TaskCreationOptions</param>
        /// <returns></returns>
        public static Task<T> StartTask<T>(Func<T> func, CancellationToken token, TaskCreationOptions options = TaskCreationOptions.None)
        {
            CheckInitialization();

            return Task.Factory.StartNew(func, token, options, UiTaskScheduler);
        }

        //This separate overload is necessary because CancellationToken.None is not a compile-time constant and thus cannot be used as default value.
        public static Task<T> StartTask<T>(Func<T> func, TaskCreationOptions options = TaskCreationOptions.None)
        {
            return StartTask(func, CancellationToken.None, options);
        }


        private static void CheckInitialization()
        {
            if (UiContext == null) throw new InvalidOperationException("UiDispatcher is not initialized. Invoke Initialize() from UI thread first.");
        }

        public static void Shutdown()
        {
            //Invoke(() =>
            //{
            //    LogManager.GetCurrentClassLogger().Debug("Invoking shutdown on UI thread dispatcher.");
            //    Dispatcher.CurrentDispatcher.InvokeShutdown();
            //});
        }
    }
}
