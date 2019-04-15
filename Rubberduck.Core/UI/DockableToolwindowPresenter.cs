using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.Resources.Registration;
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

        protected DockableToolwindowPresenter(IVBE vbe, IAddIn addin, IDockableUserControl view, IConfigurationService<WindowSettings> settingsProvider)
        {
            _vbe = vbe;
            _addin = addin;
            Logger.Trace($"Initializing Dockable Panel ({GetType().Name})");
            UserControl = view as UserControl;
            if (settingsProvider != null)
            {
                _settings = settingsProvider.Read();
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
                    var info = windows.CreateToolWindow(_addin, RubberduckProgId.DockableWindowHostProgId,
                        control.Caption, control.ClassId);
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

            toolWindow.IsVisible = true; //window resizing doesn't work without this
            EnsureMinimumWindowSize(toolWindow);
            toolWindow.IsVisible = _settings != null && _settings.IsWindowVisible(this);

            // currently we always inject _DockableToolWindowHost from Rubberduck.Main.
            // that method is not exposed in any of the interfaces we know, though, so we need to invoke it blindly
            using (var mainWindow = _vbe.MainWindow)
            {
                ((dynamic) _userControlObject).AddUserControl(control as UserControl, new IntPtr(mainWindow.HWnd));
            }

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

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            Logger.Trace($"Disposing DockableWindowPresenter of type {this.GetType()}.");

            _window.Dispose();
           
            _isDisposed = true;
        }

#if DEBUG
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1063")] // only logging here.
        ~DockableToolwindowPresenter()
        {
            // destructor for tracking purposes only - do not suppress unless 
            Debug.WriteLine($"DockableToolwindowPresenter of type {this.GetType()} finalized.");
        }
#endif
    }
}
