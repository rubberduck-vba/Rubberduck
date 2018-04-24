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
        /// from the UI thread, the action is executed immendiately. If the
        /// method is called from another thread, the action will be enqueued
        /// on the UI thread's dispatcher and executed asynchronously.
        /// <para>For additional operations on the UI thread, you can get a
        /// reference to the UI thread's context thanks to the property
        /// <see cref="UiContext" /></para>.
        /// </summary>
        /// <param name="action">The action that will be executed on the UI
        /// thread</param>
        void Invoke(Action action);

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