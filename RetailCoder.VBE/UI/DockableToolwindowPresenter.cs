using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;

namespace Rubberduck.UI
{
    public interface IPresenter
    {
        void Show();
        void Hide();
    }

    public abstract class DockableToolwindowPresenter : IPresenter, IDisposable
    {
        private readonly AddIn _addin;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly Window _window;
        protected readonly UserControl UserControl;

        protected DockableToolwindowPresenter(VBE vbe, AddIn addin, IDockableUserControl view)
        {
            _vbe = vbe;
            _addin = addin;
            Logger.Trace(string.Format("Initializing Dockable Panel ({0})", GetType().Name));
            UserControl = view as UserControl;
            _window = CreateToolWindow(view);
        }

        private readonly VBE _vbe;
        protected VBE VBE { get { return _vbe; } }

        private Window CreateToolWindow(IDockableUserControl control)
        {
            object userControlObject = null;
            Window toolWindow;
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
            toolWindow.Visible = true; //window resizing doesn't work without this

            EnsureMinimumWindowSize(toolWindow);

            toolWindow.Visible = false; //hide it again

            userControlHost.AddUserControl(control as UserControl, new IntPtr(_vbe.MainWindow.HWnd));
            return toolWindow;
        }

        private void EnsureMinimumWindowSize(Window window)
        {
            const int defaultWidth = 350;
            const int defaultHeight = 200;

            if (!window.Visible || window.LinkedWindows != null)
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
            _window.Visible = true;
        }

        public void Hide()
        {
            _window.Visible = false;
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
