using System;
using System.Runtime.InteropServices;
using System.Timers;
using EasyHook;
using Rubberduck.VBEditor.VbeRuntime;

namespace Rubberduck.Runtime
{
    public class BeepEventArgs : EventArgs
    {
        public bool Handled { get; set; }
    }

    public interface IBeepInterceptor : IDisposable
    {
        void SuppressBeep(double millisecondsToSuppress);
        event EventHandler<BeepEventArgs> Beep;
    }

    public sealed class BeepInterceptor : IBeepInterceptor
    {
        public event EventHandler<BeepEventArgs> Beep;

        private readonly IVbeNativeApi _vbeApi;
        private readonly LocalHook _hook;
        private readonly Timer _timer;

        public BeepInterceptor(IVbeNativeApi vbeApi)
        {
            _vbeApi = vbeApi;
            _hook = HookVbaBeep();
            _timer = new Timer();
            _timer.Elapsed += TimerElapsed;
        }

        public void SuppressBeep(double millisecondsToSuppress)
        {
            _suppressed = true;
            _timer.Interval = millisecondsToSuppress;
            _timer.Enabled = true;
        }

        private LocalHook HookVbaBeep()
        {
            var processAddress = LocalHook.GetProcAddress(_vbeApi.DllName, "rtcBeep");
            var callbackDelegate = new VbaBeepDelegate(VbaBeepCallback);
            var hook = LocalHook.Create(processAddress, callbackDelegate, null);
            hook.ThreadACL.SetInclusiveACL(new[] { 0 });
            return hook;
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void VbaBeepDelegate();

        private bool _suppressed;
        private void VbaBeepCallback()
        {
            if (_suppressed)
            {
                return;
            }
            
            var e = new BeepEventArgs();
            Beep?.Invoke(this, e);

            if (!e.Handled)
            {
                _vbeApi.Beep();
            }
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            _suppressed = false;
            _timer.Enabled = false;
        }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _hook?.Dispose();
            _timer?.Dispose();
            _disposed = true;
        }
    }
}
