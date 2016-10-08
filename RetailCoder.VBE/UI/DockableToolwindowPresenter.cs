using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.UI
{
    public interface IPresenter
    {
        void Show();
        void Hide();
    }

    public abstract class DockableToolwindowPresenter : IPresenter, IDisposable
    {
        private readonly IAddIn _addin;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly IWindow _window;
        protected readonly UserControl UserControl;

        protected DockableToolwindowPresenter(IVBE vbe, IAddIn addin, IDockableUserControl view)
        {
            _vbe = vbe;
            _addin = addin;
            Logger.Trace(string.Format("Initializing Dockable Panel ({0})", GetType().Name));
            UserControl = view as UserControl;
            _window = CreateToolWindow(view);
        }

        private readonly IVBE _vbe;
        protected IVBE VBE { get { return _vbe; } }

        private IWindow CreateToolWindow(IDockableUserControl control)
        {
            object userControlObject = null;
            IWindow toolWindow;
            try
            {
                toolWindow = _vbe.Windows.CreateToolWindow(_addin, _DockableWindowHost.RegisteredProgId,
                    control.Caption, control.ClassId, ref userControlObject);
            }
            catch (COMException exception)
            {
                var logEvent = new LogEventInfo(LogLevel.Error, Logger.Name, "Error Creating Control");
                logEvent.Exception = exception;
                logEvent.Properties.Add("EventID", 1);

                Logger.Error(logEvent);
                return null; //throw;
            }
            catch (NullReferenceException exception)
            {
                Logger.Error(exception);
                return null; //throw;
            }

            var userControlHost = (_DockableWindowHost)userControlObject;
            toolWindow.IsVisible = true; //window resizing doesn't work without this

            EnsureMinimumWindowSize(toolWindow);

            toolWindow.IsVisible = false; //hide it again

            userControlHost.AddUserControl(control as UserControl, new IntPtr(_vbe.MainWindow.HWnd));
            return toolWindow;
        }

        private void EnsureMinimumWindowSize(IWindow window)
        {
            const int defaultWidth = 350;
            const int defaultHeight = 200;

            if (!window.IsVisible || window.LinkedWindows != null)
            {
                return;
            }

            if (window.Width < defaultWidth)
            {
                window.Width = defaultWidth;
            }

            if (window.Height < defaultHeight)
            {
                window.Height = defaultHeight;
            }
        }

        public virtual void Show()
        {
            _window.IsVisible = true;
        }

        public void Hide()
        {
            _window.IsVisible = false;
        }

        private bool _disposed;
        public void Dispose()
        {
            Dispose(_disposed);
            _disposed = true;
        }

        public bool IsDisposed { get { return _disposed; } }
        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) { return; }
            
            if (UserControl != null)
            {
                UserControl.Dispose();
                GC.SuppressFinalize(UserControl);
            }

            if (_window != null)
            {
                _window.Close();
                Marshal.FinalReleaseComObject(_window);
            }
        }
    }
}
