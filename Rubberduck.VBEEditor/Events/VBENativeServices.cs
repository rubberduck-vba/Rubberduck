using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.VBEditor.Events
{
    public static class VBENativeServices
    {
        private static User32.WinEventProc _eventProc;
        private static IntPtr _eventHandle;
        private static IVBE _vbe;
  
        public struct WindowInfo
        {
            private readonly IntPtr _handle;
            private readonly IWindow _window;
            private readonly IWindowEventProvider _subclass;

            public IntPtr Hwnd { get { return _handle; } } 
            public IWindow Window { get { return _window; } }
            internal IWindowEventProvider Subclass { get { return _subclass; } }

            internal WindowInfo(IntPtr handle, IWindow window, IWindowEventProvider source)
            {
                _handle = handle;
                _window = window;
                _subclass = source;
            }
        }

        //This *could* be a ConcurrentDictionary, but there other operations that need the lock around it anyway.
        private static readonly Dictionary<IntPtr, WindowInfo> TrackedWindows = new Dictionary<IntPtr, WindowInfo>();
        private static readonly object ThreadLock = new object();
        
        private static uint _threadId;

        public static void HookEvents(IVBE vbe)
        {
            _vbe = vbe;
            if (_eventHandle == IntPtr.Zero)
            {               
                _eventProc = VbeEventCallback;
                _threadId = User32.GetWindowThreadProcessId(new IntPtr(_vbe.MainWindow.HWnd), IntPtr.Zero);
                _eventHandle = User32.SetWinEventHook((uint)WinEvent.Min, (uint)WinEvent.Max, IntPtr.Zero, _eventProc, 0, _threadId, WinEventFlags.OutOfContext);
            }
        }

        public static void UnhookEvents()
        {
            lock (ThreadLock)
            {
                User32.UnhookWinEvent(_eventHandle);
                foreach (var info in TrackedWindows.Values)
                {
                    info.Subclass.FocusChange -= FocusDispatcher;
                    info.Subclass.Dispose();
                }
                SafeComWrappers.VBA.VBProjects.DetatchEvents();
                SafeComWrappers.VBA.VBComponents.DetatchEvents();
            }
        }

        public static void VbeEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild,
            uint dwEventThread, uint dwmsEventTime)
        {
            if (hwnd != IntPtr.Zero && 
                idObject == (int)ObjId.Caret && 
                (eventType == (uint)WinEvent.ObjectLocationChange || eventType == (uint)WinEvent.ObjectCreate) &&
                hwnd.ToWindowType() == WindowType.CodePane)
            {
                OnSelectionChanged(hwnd);             
            }
            else if (idObject == (int)ObjId.Window && (eventType == (uint)WinEvent.ObjectCreate || eventType == (uint)WinEvent.ObjectDestroy))
            {
                var type = hwnd.ToWindowType();
                if (type != WindowType.DesignerWindow && type != WindowType.CodePane)
                {
                    return;                   
                }
                if (eventType == (uint) WinEvent.ObjectCreate)
                {
                    AttachWindow(hwnd);
                }
                else if (eventType == (uint)WinEvent.ObjectDestroy)
                {
                    DetachWindow(hwnd);
                }
            }
            else if (eventType == (uint)WinEvent.ObjectFocus && idObject == (int)ObjId.Client)
            {
                //Test to see if it was a selection change in the project window.
                var parent = User32.GetParent(hwnd);
                if (parent != IntPtr.Zero && parent.ToWindowType() == WindowType.Project && hwnd == User32.GetFocus())
                {
                    FocusDispatcher(_vbe, new WindowChangedEventArgs(parent, null, null, FocusType.ChildFocus));
                }                
            }
            //This is an output window firehose, leave this here, but comment it out when done.
            //if (idObject != (int)ObjId.Cursor) Debug.WriteLine("Hwnd: {0:X4} - EventType {1:X4}, idObject {2}, idChild {3}", (int)hwnd, eventType, idObject, idChild);
        }

        private static void AttachWindow(IntPtr hwnd)
        {
            lock (ThreadLock)
            {
                Debug.Assert(!TrackedWindows.ContainsKey(hwnd));
                var window = GetWindowFromHwnd(hwnd);
                if (window == null) return;
                var source = window.Type == WindowKind.CodeWindow
                    ? new CodePaneSubclass(hwnd, GetCodePaneFromHwnd(hwnd)) as IWindowEventProvider
                    : new DesignerWindowSubclass(hwnd);
                var info = new WindowInfo(hwnd, window, source);
                source.FocusChange += FocusDispatcher;
                TrackedWindows.Add(hwnd, info);
            }           
        }

        private static void DetachWindow(IntPtr hwnd)
        {
            lock (ThreadLock)
            {
                Debug.Assert(TrackedWindows.ContainsKey(hwnd));
                var info = TrackedWindows[hwnd];
                info.Subclass.FocusChange -= FocusDispatcher;
                info.Subclass.Dispose();
                TrackedWindows.Remove(hwnd);
            }             
        }

        private static void FocusDispatcher(object sender, WindowChangedEventArgs eventArgs)
        {
            OnWindowFocusChange(sender, eventArgs);
        }

        public static WindowInfo? GetWindowInfoFromHwnd(IntPtr hwnd)
        {
            lock (ThreadLock)
            {
                if (!TrackedWindows.ContainsKey(hwnd))
                {
                    return null;
                }
                return TrackedWindows[hwnd];
            }
        }

        public static event EventHandler<SelectionChangedEventArgs> SelectionChanged;
        private static void OnSelectionChanged(IntPtr hwnd)
        {
            if (SelectionChanged != null)
            {
                var pane = GetCodePaneFromHwnd(hwnd);
                if (pane != null) SelectionChanged.Invoke(_vbe, new SelectionChangedEventArgs(pane));
            }
        }

        public static event EventHandler<WindowChangedEventArgs> WindowFocusChange;
        private static void OnWindowFocusChange(object sender, WindowChangedEventArgs eventArgs)
        {
            if (WindowFocusChange != null)
            {
                WindowFocusChange.Invoke(sender, eventArgs);
            }
        } 

        private static ICodePane GetCodePaneFromHwnd(IntPtr hwnd)
        {
            try
            {
                var caption = hwnd.GetWindowText();
                return _vbe.CodePanes.FirstOrDefault(x => x.Window.Caption.Equals(caption));
            }
            catch
            {
                // This *should* only happen when a code pane window is removed and RD responds faster than
                // the VBE removes it from the windows collection. TODO: Find a better method to match code panes
                // to windows than testing the captions.
                return null;
            }
        }

        private static IWindow GetWindowFromHwnd(IntPtr hwnd)
        {
            if (!User32.IsWindow(hwnd))
            {
                return null;
            }

            var caption = hwnd.GetWindowText();
            return _vbe.Windows.FirstOrDefault(x => x.Caption.Equals(caption));
        }

        /// <summary>
        /// A helper function that returns <c>true</c> when the specified handle is that of the foreground window.
        /// </summary>
        /// <returns>True if the active thread is on the VBE's thread.</returns>
        public static bool IsVbeWindowActive()
        {           
            uint hThread;
            User32.GetWindowThreadProcessId(User32.GetForegroundWindow(), out hThread);
            return (IntPtr)hThread == (IntPtr)_threadId;
        }

        public enum WindowType
        {
            Indeterminate,
            VbaWindow,
            CodePane,
            DesignerWindow,
            Project
        }

        public static WindowType ToWindowType(this IntPtr hwnd)
        {
            WindowType id;
            var type = Enum.TryParse(hwnd.ToClassName(), true, out id) ? id : WindowType.Indeterminate;
            if (type != WindowType.VbaWindow)
            {
                return type;
            }
            //A this point we only care about code panes - none of the other 4 types of VbaWindows (Immediate, Object Browser, Locals,
            //and Watches) contain a tool bar at the top, so just see if the window has one as a child.
            var toolbar = User32.FindWindowEx(hwnd, IntPtr.Zero, "ObtbarWndClass", string.Empty);
            return toolbar == IntPtr.Zero ? WindowType.VbaWindow : WindowType.CodePane;
        }

        public static string ToClassName(this IntPtr hwnd)
        {
            var name = new StringBuilder(128);
            User32.GetClassName(hwnd, name, name.Capacity);
            return name.ToString();
        }
    }
}
