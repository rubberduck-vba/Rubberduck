using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.WindowsApi
{
    internal class SubclassManager : IDisposable
    {
        private static readonly Logger SubclassLogger = LogManager.GetCurrentClassLogger();
        private static readonly object ThreadLock = new object();
        private readonly ConcurrentDictionary<IntPtr, SubclassingWindow> _subclasses = new ConcurrentDictionary<IntPtr, SubclassingWindow>();

        public static bool IsSubclassable(WindowType type)
        {
            return type == WindowType.CodePane || type == WindowType.DesignerWindow;
        }

        public bool IsSubclassed(IntPtr hwnd) => _subclasses.TryGetValue(hwnd, out _);

        public IEnumerable<SubclassingWindow> Subclasses => _subclasses.Values;

        public void Subclass(IEnumerable<IntPtr> hwnds)
        {
            // ReSharper disable once ReturnValueOfPureMethodIsNotUsed (lazy coder's for-each).
            hwnds.Select(Subclass);
        }

        public SubclassingWindow Subclass(IntPtr hwnd)
        {
            var windowType = hwnd.ToWindowType();
            if (windowType == WindowType.Indeterminate)
            {
                // Not the droids we're looking for.
                return null;
            }

            lock (ThreadLock)
            { 
                if (_subclasses.TryGetValue(hwnd, out var existing))
                {
                    return existing;
                }

                // Any additional cases also need to be added to IsSubclassable above.
                switch (windowType)
                {
                    case WindowType.CodePane:
                        return TrackNewCodePane(hwnd);
                    case WindowType.DesignerWindow:
                        return TrackNewDesigner(hwnd);
                    default:
                        return null;
                }
            }
        }

        private CodePaneSubclass TrackNewCodePane(IntPtr hwnd)
        {
            var codePane = new CodePaneSubclass(hwnd, null);
            try
            {
                if (_subclasses.TryAdd(hwnd, codePane))
                {
                    codePane.ReleasingHandle += SubclassRemoved;
                    codePane.CaptionChanged += AssociateCodePane;
                    SubclassLogger.Trace($"Subclassed hWnd 0x{hwnd.ToInt64():X8} as CodePane.");
                    return codePane;
                }
            }
            catch (Exception ex)
            {
                SubclassLogger.Error(ex);
            }
            codePane.Dispose();
            return null;
        }

        private DesignerWindowSubclass TrackNewDesigner(IntPtr hwnd)
        {
            var designer = new DesignerWindowSubclass(hwnd);
            try
            {
                if (_subclasses.TryAdd(hwnd, designer))
                {
                    designer.ReleasingHandle += SubclassRemoved;
                    SubclassLogger.Trace($"Subclassed hWnd 0x{hwnd.ToInt64():X8} as DesignerWindow.");
                    return designer;
                }                
            }
            catch (Exception ex)
            {
                SubclassLogger.Error(ex);
            }
            designer.Dispose();
            return null;
        }

        private void SubclassRemoved(object sender, EventArgs eventArgs)
        {
            var subclass = (SubclassingWindow)sender;

            if (_subclasses.TryRemove(subclass.Hwnd, out _))
            {
                SubclassLogger.Trace($"Releasing subclass for hWnd 0x{subclass.Hwnd.ToInt64():X8}.");
            }
            else
            {
                SubclassLogger.Warn($"Untracked subclass with hWnd 0x{subclass.Hwnd.ToInt64():X8} released itself.");
            }
        }

        private static void AssociateCodePane(object sender, EventArgs eventArgs)
        {
            var subclass = (CodePaneSubclass)sender;
            subclass.VbeObject = VbeNativeServices.GetCodePaneFromHwnd(subclass.Hwnd);
            SubclassLogger.Trace($"CodePane subclass for hWnd 0x{subclass.Hwnd.ToInt64():X8} associated itself with its VBE object.");
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
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
                foreach (var subclass in Subclasses)
                {
                    subclass.Dispose();
                }
            }

            _disposed = true;
        }
    }
}
