using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.VBEditor.Utility;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.Parsing.UIContext
{
    public class UiDispatcher : IUiDispatcher
    {
        // thanks to Pellared on http://stackoverflow.com/a/12909070/1188513

        private readonly IUiContextProvider _contextProvider;

        public UiDispatcher(IUiContextProvider contextProvider)
        {
            _contextProvider = contextProvider;
        }

        /// <inheritdoc />
        public void InvokeAsync(Action action)
        {
            CheckInitialization();

            _contextProvider.UiContext.Post(x => action(), null);
        }

        /// <inheritdoc />
        public void Invoke(Action action)
        {
            CheckInitialization();

            if (_contextProvider.UiContext == SynchronizationContext.Current)
            {
                action();
            }
            else
            {
                InvokeAsync(action);
            }
        }

        /// <inheritdoc />
        public void FlushMessageQueue()
        {
            CheckInitialization();

            if (_contextProvider.UiContext == SynchronizationContext.Current)
            {
                PumpMessages();
            }
            else
            {
                InvokeAsync(PumpMessages);
            }

            // This should remain a local function - messages should not be pumped out outside of FlushMessageQueue.
            void PumpMessages()
            {
                var message = new NativeMethods.NativeMessage();
                var handle = GCHandle.Alloc(message);

                while (NativeMethods.PeekMessage(ref message, IntPtr.Zero, 0, 0, NativeMethods.PeekMessageRemoval.Remove))
                {
                    NativeMethods.TranslateMessage(ref message);
                    NativeMethods.DispatchMessage(ref message);
                }

                handle.Free();
            }
        }

        private const uint RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;
        private const uint VBA_E_IGNORE = 0x800AC472;
        private const uint VBA_E_CANTEXECCODEINBREAKMODE = 0x800ADF09;

        /// <inheritdoc />
        public void RaiseComEvent(Action comEventHandler)
        {
            Invoke(() =>
            {
                var currentCount = 0;
                var retryCount = 100;
                var timeSleep = 10;
                for (; ; )
                {
                    try
                    {
                        comEventHandler.Invoke();
                        break;
                    }
                    catch (Exception ex)
                    {
                        if (currentCount < retryCount)
                        {
                            var cex = (COMException)ex;
                            switch ((uint)cex.ErrorCode)
                            {
                                case VBA_E_CANTEXECCODEINBREAKMODE:
                                    Thread.Sleep(timeSleep);
                                    break;
                                case RPC_E_SERVERCALL_RETRYLATER:
                                    Thread.Sleep(timeSleep);
                                    break;
                                case VBA_E_IGNORE:
                                    Thread.Sleep(timeSleep);
                                    break;
                                default:
                                    throw;
                            }

                        }
                        else
                        {
                            throw;
                        }
                        currentCount++;
                    }
                }
            });
        }

        /// <inheritdoc />
        public Task StartTask(Action action, CancellationToken token, TaskCreationOptions options = TaskCreationOptions.None)
        {
            CheckInitialization();

            if (_contextProvider.UiContext == SynchronizationContext.Current)
            {
                action.Invoke();
                return Task.CompletedTask;
            }

            return Task.Factory.StartNew(action, token, options, _contextProvider.UiTaskScheduler);
        }

        //This separate overload is necessary because CancellationToken.None is not a compile-time constant and thus cannot be used as default value.
        public Task StartTask(Action action, TaskCreationOptions options = TaskCreationOptions.None)
        {
            return StartTask(action, CancellationToken.None, options);
        }


        /// <inheritdoc />
        public Task<T> StartTask<T>(Func<T> func, CancellationToken token, TaskCreationOptions options = TaskCreationOptions.None)
        {
            CheckInitialization();

            if (_contextProvider.UiContext == SynchronizationContext.Current)
            {
                var returnValue = func();
                return Task.FromResult(returnValue);
            }

            return Task.Factory.StartNew(func, token, options, _contextProvider.UiTaskScheduler);
        }

        //This separate overload is necessary because CancellationToken.None is not a compile-time constant and thus cannot be used as default value.
        public Task<T> StartTask<T>(Func<T> func, TaskCreationOptions options = TaskCreationOptions.None)
        {
            return StartTask(func, CancellationToken.None, options);
        }

        /// <remarks>
        /// Depends on the static method: <see cref="UiContextProvider.Initialize"/>
        /// </remarks>
        private void CheckInitialization()
        {
            if (_contextProvider.UiContext == null)
            {
                throw new InvalidOperationException("UiContext is not initialized. Invoke UiContextProvider.Initialize() from the UI thread first.");
            }
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
