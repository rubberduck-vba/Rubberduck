using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.UI
{
    public abstract class SubclassingWindow : IDisposable
    {
        private readonly IntPtr _subclassId;
        private readonly IntPtr _hwnd;
        private readonly SubClassCallback _wndProc;
        private bool _listening;

        private static readonly ConcurrentBag<SubClassCallback> RubberduckProcs = new ConcurrentBag<SubClassCallback>();
        private static readonly object SubclassLock = new object();

        public delegate int SubClassCallback(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int SetWindowSubclass(IntPtr hWnd, SubClassCallback newProc, IntPtr uIdSubclass, IntPtr dwRefData);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int RemoveWindowSubclass(IntPtr hWnd, SubClassCallback newProc, IntPtr uIdSubclass);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int DefSubclassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam);

        public IntPtr Hwnd { get; set; }

        protected SubclassingWindow(IntPtr subclassId, IntPtr hWnd)
        {
            _subclassId = subclassId;
            _hwnd = hWnd;
            _wndProc = SubClassProc;
            AssignHandle();
        }
        ~SubclassingWindow()
        {
            Debug.Assert(false, "Dispose() not called.");
        }

        public void Dispose()
        {
            ReleaseHandle();
            GC.SuppressFinalize(this);
        }

        private void AssignHandle()
        {
            lock (SubclassLock)
            {
                RubberduckProcs.Add(_wndProc);
                var result = SetWindowSubclass(_hwnd, _wndProc, _subclassId, IntPtr.Zero);
                if (result != 1)
                {
                    throw new Exception("SetWindowSubClass Failed");
                }
                _listening = true;
            }
        }

        private void ReleaseHandle()
        {
            if (!_listening)
            {
                return;
            }

            lock (SubclassLock)
            {
                var result = RemoveWindowSubclass(_hwnd, _wndProc, _subclassId);
                if (result != 1)
                {
                    throw new Exception("RemoveWindowSubclass Failed");
                }
                _listening = false;
            }
        }

        public virtual int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            if (!_listening)
            {
                throw new Exception("State corrupted. Received window message while not listening.");
            }

            Debug.Assert(IsWindow(_hwnd));
            if ((uint)msg == (uint)WM.RUBBERDUCK_SINKING)
            {
                ReleaseHandle();
            }
            return DefSubclassProc(hWnd, msg, wParam, lParam);
        }
    }
}