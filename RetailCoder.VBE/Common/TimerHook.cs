using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Rubberduck.Common.WinAPI;
using NLog;

namespace Rubberduck.Common
{
    public class TimerHook : IAttachable, IDisposable
    {
        private readonly IntPtr _mainWindowHandle;
        private readonly User32.TimerProc _timerProc;
        private readonly Logger _log = LogManager.GetCurrentClassLogger();
        private IntPtr _timerId;

        public TimerHook(IntPtr mainWindowHandle)
        {
            _mainWindowHandle = mainWindowHandle;
            _timerProc = TimerCallback;
        }

        public bool IsAttached => _timerId != IntPtr.Zero;

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
                if (timerId == IntPtr.Zero)
                {
                    _log.Warn($"Unable to create a global atom for timer hook; error was {Marshal.GetLastWin32Error()}; aborting the Attach operation...");
                    return;
                }

                if (User32.SetTimer(_mainWindowHandle, timerId, 500, _timerProc) != IntPtr.Zero)
                {
                    _timerId = timerId;
                }
                else
                {
                    _log.Warn($"Global atom was created but SetTime failed with error {Marshal.GetLastWin32Error()}; aborting the Attach operation...");
                    Kernel32.SetLastError(Kernel32.ERROR_SUCCESS);
                    Kernel32.GlobalDeleteAtom(timerId);
                    var lastError = Marshal.GetLastWin32Error();
                    if(lastError!=Kernel32.ERROR_SUCCESS)
                    {
                        _log.Warn($"Deleting global atom failed with error {lastError}; the atom was {timerId}.");
                    };
                }
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
                if (!User32.KillTimer(_mainWindowHandle, _timerId))
                {
                    _log.Warn($"Error with executing KillTimer; the error was {Marshal.GetLastWin32Error()}; continuing with deletion of atom.");
                }

                Kernel32.SetLastError(Kernel32.ERROR_SUCCESS);
                Kernel32.GlobalDeleteAtom(_timerId);
                var lastError = Marshal.GetLastWin32Error();
                if(lastError!=Kernel32.ERROR_SUCCESS)
                {
                    _log.Warn($"Error with executing GlobalDeleteAtom; the error was {lastError}; the hook will be deatched anyway; the timerId was {_timerId}");
                }

                _timerId = IntPtr.Zero;
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
