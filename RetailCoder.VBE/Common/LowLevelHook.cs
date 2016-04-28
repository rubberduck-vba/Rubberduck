using Rubberduck.Common.WinAPI;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace Rubberduck.Common
{
    public abstract class LowLevelHook : IAttachable
    {
        private Thread _thread;
        private readonly User32.HookProc _hookProc;
        private ApplicationContext _context;
        private readonly WindowsHook _hookType;
        private IntPtr _hookId;
        private bool _isAttached;

        public LowLevelHook(WindowsHook hookType)
        {
            _hookType = hookType;
            _hookProc = new User32.HookProc(HookCallback);
        }

        public bool IsAttached
        {
            get
            {
                return _isAttached;
            }
        }

        public void Attach()
        {
            if (_isAttached)
            {
                Detach();
            }
            _isAttached = true;
            _context = new ApplicationContext();
            _thread = new Thread(Hook);
            _thread.Start();
        }

        public void Detach()
        {
            try
            {
                Debug.WriteLine("Stopping hook {0}.", _hookType);
                Debug.WriteLine("Stopping message pump.");
                _context.ExitThread();
                Debug.WriteLine("Message pump gone. Stopping thread.");
                _thread.Join();
                Debug.WriteLine("Thread gone. Unhooking hook.");
                Unhook();
                Debug.WriteLine("Unhooked. All done.");
            }
            catch
            {
                // Shutting down the process would forcefully clean the resources.
            }
            finally
            {
                _isAttached = false;
            }
        }

        private void Hook()
        {
            var handle = Kernel32.GetModuleHandle("user32");
            if (handle == IntPtr.Zero)
            {
                throw new Win32Exception();
            }
            _hookId = User32.SetWindowsHookEx(_hookType, _hookProc, handle, 0);
            if (_hookId == IntPtr.Zero)
            {
                throw new Win32Exception();
            }
            Application.Run(_context);
            Debug.WriteLine("Message pump for hook {0} has stopped.", _hookType);
        }

        public void Unhook()
        {
            try
            {
                User32.UnhookWindowsHookEx(_hookId);
            }
            finally
            {
                _hookId = IntPtr.Zero;
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return User32.CallNextHookEx(_hookId, nCode, wParam, lParam);
            }
            try
            {
                HookCallbackCore(nCode, wParam, lParam);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            return User32.CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        protected abstract void HookCallbackCore(int nCode, IntPtr wParam, IntPtr lParam);

        public event EventHandler<HookEventArgs> MessageReceived;
        protected void OnMessageReceived()
        {
            var handler = MessageReceived;
            if (handler != null)
            {
                handler.Invoke(this, HookEventArgs.Empty);
            }
        }
    }
}
