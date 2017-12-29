using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI
{
    public interface IPresenter
    {
        void Show();
        void Hide();
    }

    public interface IDockablePresenter : IPresenter
    {
        UserControl UserControl { get; }
    }

    public abstract class DockableToolwindowPresenter : IDockablePresenter, IDisposable
    {
        private readonly IAddIn _addin;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly IWindow _window;
        private readonly WindowSettings _settings;  //Storing this really doesn't matter - it's only checked on startup and never persisted.

        protected DockableToolwindowPresenter(IVBE vbe, IAddIn addin, IDockableUserControl view, IConfigProvider<WindowSettings> settingsProvider)
        {
            _vbe = vbe;
            _addin = addin;
            Logger.Trace($"Initializing Dockable Panel ({GetType().Name})");
            UserControl = view as UserControl;
            if (settingsProvider != null)
            {
                _settings = settingsProvider.Create();
            }
            _window = CreateToolWindow(view);
        }

        public UserControl UserControl { get; }

        private object _userControlObject;
        private readonly IVBE _vbe;

        private IWindow CreateToolWindow(IDockableUserControl control)
        {
            IWindow toolWindow;
            try
            {
                using (var windows = _vbe.Windows)
                {
                    var info = windows.CreateToolWindow(_addin, _DockableWindowHost.RegisteredProgId, control.Caption, control.ClassId);
                    _userControlObject = info.UserControl;
                    toolWindow = info.ToolWindow;
                }
            }
            catch (COMException exception)
            {
                Logger.Error(exception);
                throw;
            }
            catch (NullReferenceException exception)
            {
                Logger.Error(exception);
                throw;
            }

            var userControlHost = (_DockableWindowHost)_userControlObject;
            toolWindow.IsVisible = true; //window resizing doesn't work without this

            EnsureMinimumWindowSize(toolWindow);

            toolWindow.IsVisible = _settings != null && _settings.IsWindowVisible(this);

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

        public virtual void Show() => _window.IsVisible = true;
        public virtual void Hide() => _window.IsVisible = false;


        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            Logger.Trace($"Disposing DockableWindowPresenter of type {this.GetType()}.");

            _window.Dispose();

            _isDisposed = true;
        }


        ~DockableToolwindowPresenter()
        {
            // destructor for tracking purposes only - do not suppress unless 
            Debug.WriteLine($"DockableToolwindowPresenter of type {this.GetType()} finalized.");
        }
    }
}
