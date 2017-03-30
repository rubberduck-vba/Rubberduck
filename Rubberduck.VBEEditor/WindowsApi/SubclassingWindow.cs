using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.WindowsApi
{
    public abstract class SubclassingWindow : IDisposable
    {
        private readonly IntPtr _subclassId;
        private readonly IntPtr _hwnd;
        private readonly SubClassCallback _wndProc;
        private bool _listening;
        private GCHandle _thisHandle;

        private readonly object _subclassLock = new object();

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

        public IntPtr Hwnd { get { return _hwnd; } }

        protected SubclassingWindow(IntPtr subclassId, IntPtr hWnd)
        {
            _subclassId = subclassId;
            _hwnd = hWnd;
            _wndProc = SubClassProc;
            AssignHandle();
        }

        public void Dispose()
        {
            ReleaseHandle();
            _thisHandle.Free();
        }

        private void AssignHandle()
        {
            lock (_subclassLock)
            {
                var result = SetWindowSubclass(_hwnd, _wndProc, _subclassId, IntPtr.Zero);
                if (result != 1)
                {
                    throw new Exception("SetWindowSubClass Failed");
                }
                Debug.WriteLine("SubclassingWindow.AssignHandle called for hWnd " + Hwnd);
                //DO NOT REMOVE THIS CALL. Dockable windows are instantiated by the VBE, not directly by RD.  On top of that,
                //since we have to inherit from UserControl we don't have to keep handling window messages until the VBE gets
                //around to destroying the control's host or it results in an access violation when the base class is disposed.
                //We need to manually call base.Dispose() ONLY in response to a WM_DESTROY message.
                _thisHandle = GCHandle.Alloc(this, GCHandleType.Normal);
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
                Debug.WriteLine("State corrupted. Received window message while not listening.");
                return DefSubclassProc(hWnd, msg, wParam, lParam);
            }

            if ((uint)msg == (uint)WM.RUBBERDUCK_SINKING || (uint)msg == (uint)WM.DESTROY)
            {               
                ReleaseHandle();                
            }
            return DefSubclassProc(hWnd, msg, wParam, lParam);
        }
    }
}