using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using NLog;

namespace Rubberduck.UI
{
    public abstract class DockablePresenterBase : IDisposable
    {
        private readonly AddIn _addin;
        private readonly Logger _logger;
        private readonly Window _window;
        protected readonly UserControl UserControl;

        protected DockablePresenterBase(VBE vbe, AddIn addin, IDockableUserControl control)
        {
            _vbe = vbe;
            _addin = addin;
            _logger = LogManager.GetCurrentClassLogger();
            _logger.Trace("Initializing Dockable Panel");
            UserControl = control as UserControl;
            _window = CreateToolWindow(control);
        }

        private readonly VBE _vbe;
        protected VBE VBE { get { return _vbe; } }

        private Window CreateToolWindow(IDockableUserControl control)
        {
            object userControlObject = null;
            Window toolWindow;
            try
            {
                _logger.Trace("Loading \"{0}\" ClassId {1}", control.Caption, control.ClassId);
                toolWindow = _vbe.Windows.CreateToolWindow(_addin, _DockableWindowHost.RegisteredProgId,
                    control.Caption, control.ClassId, ref userControlObject);
            }
            catch (COMException exception)
            {
                var logEvent = new LogEventInfo(LogLevel.Error, _logger.Name,  "Error Creating Control");
                logEvent.Exception = exception;
                logEvent.Properties.Add("EventID", 1);

                _logger.Error(logEvent);
                throw;
            }
            var userControlHost = (_DockableWindowHost)userControlObject;
            toolWindow.Visible = true; //window resizing doesn't work without this

            EnsureMinimumWindowSize(toolWindow);

            toolWindow.Visible = false; //hide it again

            userControlHost.AddUserControl(control as UserControl);
            return toolWindow;
        }

        private void EnsureMinimumWindowSize(Window window)
        {
            const int defaultWidth = 350;
            const int defaultHeight = 200;

            if (window.Visible && window.LinkedWindows == null) //checking these conditions prevents errors
            {
                if (window.Width < defaultWidth)
                {
                    window.Width = defaultWidth;
                }

                if (window.Height < defaultHeight)
                {
                    window.Height = defaultHeight;
                }
            }
        }

        public virtual void Show()
        {
            _window.Visible = true;
        }

        public virtual void Close()
        {
            _window.Close();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) { return; }
            
            if (UserControl != null)
            {
                UserControl.Dispose();
            }

            Marshal.ReleaseComObject(_window);
        }
    }
}