using Rubberduck.VBEditor.Utility;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.UIContext
{
    public interface IUiDispatcher
    {
        /// <summary>
        /// Invokes an action asynchronously on the UI thread.
        /// </summary>
        /// <param name="action">The action that must be executed.</param>
        void InvokeAsync(Action action);

        /// <summary>
        /// Executes an action on the UI thread. If this method is called
        /// from the UI thread, the action is executed immediately. If the
        /// method is called from another thread, the action will be enqueued
        /// on the UI thread's dispatcher and executed asynchronously.
        /// <para>For additional operations on the UI thread, you can get a
        /// reference to the UI thread's context thanks to the property
        /// <see cref="UiContextProvider" /></para>.
        /// </summary>
        /// <param name="action">The action that will be executed on the UI
        /// thread</param>
        void Invoke(Action action);

        /// <summary>
        /// Flushes all pending messages on the UI thread. If called from the UI thread, it will immediately yield to
        /// any other UI activity ala DoEvents. If called from a non-UI thread, it will queue on the UI thread (and then
        /// basically do nothing because... it should be the only thing left at that point) - the capability for this to
        /// be "asynchronously executed" is only provided to ensure it is never executed out of context.
        /// </summary>
        void FlushMessageQueue();

        /// <summary>
        /// Raises a COM-visible event on the UI thread. This will use <see cref="UiDispatcher.Invoke" /> internally
        /// but with additional error handling and retry logic for transient failure to fire COM event due to the host
        /// being too busy to accept event.
        /// </summary>
        /// <param name="comEventHandler">The handler for setting up and firing the COM event on the UI thread</param>
        void RaiseComEvent(Action comEventHandler);

        /// <summary>
        /// Starts a task on the ui thread.
        /// </summary>
        /// <param name="action">The action that will be executed on the UI
        /// thread.</param>
        /// <param name="token">Optional cancellation token</param>
        /// <param name="options">Optional TaskCreationOptions</param>
        /// <returns></returns>
        Task StartTask(Action action, CancellationToken token, TaskCreationOptions options = TaskCreationOptions.None);

        Task StartTask(Action action, TaskCreationOptions options = TaskCreationOptions.None);

        /// <summary>
        /// Starts a task returning a value on the ui thread.
        /// </summary>
        /// <param name="func">The function that will be executed on the UI
        /// thread.</param>
        /// <param name="token">Optional cancellation token</param>
        /// <param name="options">Optional TaskCreationOptions</param>
        /// <returns></returns>
        Task<T> StartTask<T>(Func<T> func, CancellationToken token, TaskCreationOptions options = TaskCreationOptions.None);

        Task<T> StartTask<T>(Func<T> func, TaskCreationOptions options = TaskCreationOptions.None);
    }
}