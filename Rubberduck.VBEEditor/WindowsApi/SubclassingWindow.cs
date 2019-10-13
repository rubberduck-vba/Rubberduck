using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using NLog;

namespace Rubberduck.VBEditor.WindowsApi
{
    public abstract class SubclassingWindow : IDisposable
    {
        protected static readonly Logger SubclassLogger = LogManager.GetCurrentClassLogger();
        public event EventHandler ReleasingHandle;
        private readonly IntPtr _subclassId;
        private readonly SubClassCallback _wndProc;
        private bool _listening;

        private readonly object _subclassLock = new object();

        public delegate int SubClassCallback(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int SetWindowSubclass(IntPtr hWnd, SubClassCallback newProc, IntPtr uIdSubclass, IntPtr dwRefData);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int RemoveWindowSubclass(IntPtr hWnd, SubClassCallback newProc, IntPtr uIdSubclass);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int DefSubclassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam);

        public IntPtr Hwnd { get; }

        protected SubclassingWindow(IntPtr subclassId, IntPtr hWnd)
        {
            _subclassId = subclassId;
            _wndProc = SubClassProc;
            Hwnd = hWnd;

            AssignHandle();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void AssignHandle()
        {
            lock (_subclassLock)
            {
                var result = SetWindowSubclass(Hwnd, _wndProc, _subclassId, IntPtr.Zero);
                if (result != 1)
                {
                    throw new Exception("SetWindowSubClass Failed");
                }

                _listening = true;
            }
        }

        private void ReleaseHandle()
        {
            lock (_subclassLock)
            {
                if (!_listening)
                {
                    return;
                }
                Debug.WriteLine("SubclassingWindow.ReleaseHandle called for hWnd " + Hwnd);
                var result = RemoveWindowSubclass(Hwnd, _wndProc, _subclassId);
                if (result != 1)
                {
                    throw new Exception("RemoveWindowSubclass Failed");
                }
                ReleasingHandle?.Invoke(this, null);
                ReleasingHandle = delegate { };
                _listening = false;
            }
        }

        public virtual int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            if (!_listening)
            {
                Debug.WriteLine("State corrupted. Received window message while not listening.");
                return DefSubclassProc(hWnd, msg, wParam, lParam);
            }

            PeekMessagePump(hWnd, msg, wParam, lParam);

            if ((uint) msg == (uint) WM.DESTROY)
            {
                Dispose();
            }
            return DefSubclassProc(hWnd, msg, wParam, lParam);
        }

        [Conditional("THIRSTY_DUCK")]
        [Conditional("THIRSTY_DUCK_WM")]
        private static void PeekMessagePump(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam)
        {
            //This is an output window firehose kind of like spy++. Prepare for some spam.
            var windowClassName = ToClassName(hWnd);
            if (WM_MAP.Lookup.TryGetValue((uint) msg, out var wmName))
            {
                wmName = $"WM_{wmName}";
            }
            else
            {
                wmName = "Unknown";
            }


            Debug.WriteLine(
                $"MSG: 0x{(uint) msg:X4} ({wmName}), Hwnd 0x{(uint) hWnd:X4} ({windowClassName}), wParam 0x{(uint) wParam:X4}, lParam 0x{(uint) lParam:X4}");
        }

        private bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                ReleaseHandle();
            }

            _disposed = true;
        }

        private static string ToClassName(IntPtr hwnd)
        {
            var name = new StringBuilder(User32.MaxGetClassNameBufferSize);
            User32.GetClassName(hwnd, name, name.Capacity);
            return name.ToString();
        }
    }
}