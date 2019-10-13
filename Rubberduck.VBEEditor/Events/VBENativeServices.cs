using System;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.VBEditor.Events
{
    // ReSharper disable once InconsistentNaming
    public static class VbeNativeServices
    {
        private static User32.WinEventProc _eventProc;
        private static IntPtr _eventHandle;
        private static IVBE _vbe;
        private static readonly SubclassManager Subclasses = new SubclassManager(); 
        private static readonly object ThreadLock = new object();        
        private static uint _threadId;

        public static void HookEvents(IVBE vbe)
        {
            _vbe = vbe;
            if (_eventHandle == IntPtr.Zero)
            {               
                _eventProc = VbeEventCallback;
                IntPtr mainWindowHwnd;
                using (var mainWindow = _vbe.MainWindow)
                {
                    mainWindowHwnd = new IntPtr(mainWindow.HWnd);
                }
                _threadId = User32.GetWindowThreadProcessId(mainWindowHwnd, IntPtr.Zero);
                _eventHandle = User32.SetWinEventHook((uint)WinEvent.Min, (uint)WinEvent.Max, IntPtr.Zero, _eventProc, 0, _threadId, WinEventFlags.OutOfContext);

                Subclasses.Subclass(mainWindowHwnd.ChildWindows()
                    .Where(hwnd => SubclassManager.IsSubclassable(hwnd.ToWindowType())));
            }
        }

        public static void UnhookEvents()
        {
            lock (ThreadLock)
            {
                SelectionChanged = delegate { };
                IntelliSenseChanged = delegate { };
                KeyDown = delegate { };
                WindowFocusChange = delegate { };
                User32.UnhookWinEvent(_eventHandle);
                Subclasses.Dispose();
                VbeEvents.Terminate();
                _vbe = null;
            }
        }

        private static bool Suspend { get; set; }

        private static void Attach(IntPtr hwnd)
        {
            var subclass = Subclasses.Subclass(hwnd);

            if (subclass == null)
            {
                return;
            }

            if (subclass is IFocusProvider focusSource)
            {
                focusSource.FocusChange += FocusDispatcher;
            }

            if (subclass is IWindowEventProvider keyboardListener)
            {
                keyboardListener.KeyDown += KeyDownDispatcher;
            }
        }

        public static void VbeEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild,
            uint dwEventThread, uint dwmsEventTime)
        {
            if (Suspend || hwnd == IntPtr.Zero || _vbe.IsWrappingNullReference) { return; }

            var windowType = hwnd.ToWindowType();

            PeekMessagePump(eventType, hwnd, idObject, idChild);

            if (windowType == WindowType.IntelliSense)
            {
                if (eventType == (uint)WinEvent.ObjectShow)
                {
                    OnIntelliSenseChanged(true);
                }
                else if (eventType == (uint)WinEvent.ObjectHide)
                {
                    OnIntelliSenseChanged(false);
                }
            }
            else if (windowType == WindowType.CodePane && idObject == (int)ObjId.Caret && 
                (eventType == (uint)WinEvent.ObjectLocationChange || eventType == (uint)WinEvent.ObjectCreate))
            {
                OnSelectionChanged(hwnd);
            }
            else if (SubclassManager.IsSubclassable(windowType) && (idObject == (int)ObjId.Window && eventType == (uint)WinEvent.ObjectCreate) ||
                     !Subclasses.IsSubclassed(hwnd))
            {
                Attach(hwnd);
            }
            else if (eventType == (uint)WinEvent.ObjectFocus && idObject == (int)ObjId.Client)
            {
                //Test to see if it was a selection change in the project window.
                var parent = User32.GetParent(hwnd);
                if (parent != IntPtr.Zero && parent.ToWindowType() == WindowType.Project && hwnd == User32.GetFocus())
                {
                    FocusDispatcher(_vbe, new WindowChangedEventArgs(parent, FocusType.ChildFocus));
                }                
            }
        }

        private static void KeyDownDispatcher(object sender, KeyPressEventArgs e)
        {
             OnKeyDown(e);
        }

        private static void FocusDispatcher(object sender, WindowChangedEventArgs eventArgs)
        {
            OnWindowFocusChange(sender, eventArgs);
        }

        public static event EventHandler<SelectionChangedEventArgs> SelectionChanged;
        private static void OnSelectionChanged(IntPtr hwnd)
        {
            using (var pane = GetCodePaneFromHwnd(hwnd))
            {
                if (pane != null)
                {
                    SelectionChanged?.Invoke(_vbe, new SelectionChangedEventArgs());
                }
            }
        }

        public static event EventHandler<IntelliSenseEventArgs> IntelliSenseChanged;

        public static void OnIntelliSenseChanged(bool shown)
        {
            IntelliSenseChanged?.Invoke(_vbe, shown ? IntelliSenseEventArgs.Shown : IntelliSenseEventArgs.Hidden);
        }

        public static event EventHandler<AutoCompleteEventArgs> KeyDown;
        private static void OnKeyDown(KeyPressEventArgs e)
        {
            using (var pane = GetCodePaneFromHwnd(e.Hwnd))
            {
                if (pane != null)
                {
                    using (var module = pane.CodeModule)
                    {
                        var args = new AutoCompleteEventArgs(module, e);
                        
                        Suspend = true;
                        KeyDown?.Invoke(_vbe, args);
                        Suspend = false;
                        e.Handled = args.Handled;
                    }
                }
            }
        }

        public static event EventHandler<WindowChangedEventArgs> WindowFocusChange;
        private static void OnWindowFocusChange(object sender, WindowChangedEventArgs eventArgs)
        {
            WindowFocusChange?.Invoke(sender, eventArgs);
        } 

        public static ICodePane GetCodePaneFromHwnd(IntPtr hwnd)
        {
            if (_vbe == null || _vbe.IsWrappingNullReference)
            {
                return null;
            }

            try
            {
                var caption = hwnd.GetWindowText();
                using (var panes = _vbe.CodePanes)
                {
                    if (panes == null || panes.IsWrappingNullReference)
                    {
                        return null;
                    }

                    var foundIt = false;
                    foreach (var pane in panes)
                    {
                        try
                        {
                            using (var window = pane.Window)
                            {
                                if (window.Caption.Equals(caption))
                                {
                                    foundIt = true;
                                    return pane;
                                }
                            }
                        }
                        finally
                        {
                            if (!foundIt)
                            {
                                pane.Dispose();
                            }
                        }
                    }

                    return null;
                }
            }
            catch
            {
                // This *should* only happen when a code pane window is removed and RD responds faster than
                // the VBE removes it from the windows collection. TODO: Find a better method to match code panes
                // to windows than testing the captions.
                return null;
            }
        }

        public static IWindow GetWindowFromHwnd(IntPtr hwnd)
        {
            if (!User32.IsWindow(hwnd) || _vbe == null || _vbe.IsWrappingNullReference)
            {
                return null;
            }

            var caption = hwnd.GetWindowText();
            using (var windows = _vbe.Windows)
            {
                if (windows == null || windows.IsWrappingNullReference)
                {
                    return null;
                }

                var foundIt = false;
                foreach (var window in windows)
                {
                    try
                    {
                        if (window.Caption.Equals(caption))
                        {
                            foundIt = true;
                            return window;
                        }
                    }
                    finally
                    {
                        if (!foundIt)
                        {
                            window.Dispose();
                        }
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// A helper function that returns <c>true</c> when the specified handle is that of the foreground window.
        /// </summary>
        /// <returns>True if the active thread is on the VBE's thread.</returns>
        public static bool IsVbeWindowActive()
        {
            User32.GetWindowThreadProcessId(User32.GetForegroundWindow(), out var hThread);
            return (IntPtr)hThread == (IntPtr)_threadId;
        }

        public static WindowType ToWindowType(this IntPtr hwnd)
        {
            var className = hwnd.ToClassName();
            if (className.Equals("NameListWndClass"))
            {
                return WindowType.IntelliSense;
            }

            var type = Enum.TryParse(className, true, out WindowType id) ? id : WindowType.Indeterminate;
            if (type != WindowType.VbaWindow)
            {
                return type;
            }
            //A this point we only care about code panes - none of the other 4 types of VbaWindows (Immediate, Object Browser, Locals,
            //and Watches) contain a tool bar at the top, so just see if the window has one as a child.
            var toolbar = User32.FindWindowEx(hwnd, IntPtr.Zero, "ObtbarWndClass", string.Empty);
            return toolbar == IntPtr.Zero ? WindowType.VbaWindow : WindowType.CodePane;
        }

        private static string ToClassName(this IntPtr hwnd)
        {
            var name = new StringBuilder(User32.MaxGetClassNameBufferSize);
            User32.GetClassName(hwnd, name, name.Capacity);
            return name.ToString();
        }

        [Conditional("THIRSTY_DUCK")]
        [Conditional("THIRSTY_DUCK_EVT")]
        private static void PeekMessagePump(uint eventType, IntPtr hwnd, int idObject, int idChild)
        {
            //This is an output window firehose, viewer discretion is advised.
            if (idObject != (int)ObjId.Cursor)
            {
                var windowClassName = hwnd.ToClassName();
                if (!WinEventMap.Lookup.TryGetValue((int)eventType, out var eventName))
                {
                    eventName = "Unknown";
                }
                Debug.WriteLine($"EVT: 0x{eventType:X4} ({eventName}) Hwnd 0x{(int)hwnd:X4} ({windowClassName}), idObject {idObject}, idChild {idChild}");
            }
        }
    }
}
