using System;
using System.Diagnostics;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public class TimerHook : IAttachable, IDisposable
    {
        private readonly IntPtr _mainWindowHandle;
        private readonly User32.TimerProc _timerProc;

        private IntPtr _timerId;

        public TimerHook(IntPtr mainWindowHandle)
        {
            _mainWindowHandle = mainWindowHandle;
            _timerProc = TimerCallback;
        }

        public bool IsAttached { get; private set; }

        public event EventHandler<HookEventArgs> MessageReceived;

        public void Attach()
        {
            if (IsAttached)
            {
                return;
            }

            try
            {
                var timerId = (IntPtr)Kernel32.GlobalAddAtom(Guid.NewGuid().ToString());
                User32.SetTimer(_mainWindowHandle, timerId, 500, _timerProc);
                IsAttached = true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        public void Detach()
        {
            if (!IsAttached)
            {
                Debug.Assert(_timerId == IntPtr.Zero);
                return;
            }

            try
            {
                User32.KillTimer(_mainWindowHandle, _timerId);
                Kernel32.GlobalDeleteAtom(_timerId);

                _timerId = IntPtr.Zero;
                IsAttached = false;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        private void OnTick()
        {
            MessageReceived?.Invoke(this, HookEventArgs.Empty);
        }

        private void TimerCallback(IntPtr hWnd, WindowLongFlags msg, IntPtr timerId, uint time)
        {
            OnTick();
        }

        public void Dispose()
        {
            if (IsAttached)
            {
                Detach();
            }

            Debug.Assert(_timerId == IntPtr.Zero);
        }
    }
}
