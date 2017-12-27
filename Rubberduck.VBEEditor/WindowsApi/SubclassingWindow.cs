using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.WindowsApi
{
    public abstract class SubclassingWindow : IDisposable
    {
        private readonly IntPtr _subclassId;
        private readonly SubClassCallback _wndProc;
        private bool _listening;

        private readonly object _subclassLock = new object();

        public delegate int SubClassCallback(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass,
            IntPtr dwRefData);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int SetWindowSubclass(IntPtr hWnd, SubClassCallback newProc, IntPtr uIdSubclass,
            IntPtr dwRefData);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int RemoveWindowSubclass(IntPtr hWnd, SubClassCallback newProc, IntPtr uIdSubclass);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int DefSubclassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam);

        public IntPtr Hwnd { get; }

        protected SubclassingWindow(IntPtr subclassId, IntPtr hWnd)
        {
            _subclassId = subclassId;
            Hwnd = hWnd;
            _wndProc = SubClassProc;
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
                _listening = false;
            }
        }

        public virtual int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass,
            IntPtr dwRefData)
        {
            if (!_listening)
            {
                Debug.WriteLine("State corrupted. Received window message while not listening.");
                return DefSubclassProc(hWnd, msg, wParam, lParam);
            }

            if ((uint) msg == (uint) WM.DESTROY)
            {
                Dispose();
            }
            return DefSubclassProc(hWnd, msg, wParam, lParam);
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
    }
}