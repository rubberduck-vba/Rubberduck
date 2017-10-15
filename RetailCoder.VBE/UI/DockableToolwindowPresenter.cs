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

    public abstract class DockableToolwindowPresenter : IDockablePresenter
    {
        private readonly IAddIn _addin;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly IWindow _window;
        private readonly UserControl _userControl;
        private readonly WindowSettings _settings;  //Storing this really doesn't matter - it's only checked on startup and never persisted.

        protected DockableToolwindowPresenter(IVBE vbe, IAddIn addin, IDockableUserControl view, IConfigProvider<WindowSettings> settingsProvider)
        {
            _vbe = vbe;
            _addin = addin;
            Logger.Trace(string.Format("Initializing Dockable Panel ({0})", GetType().Name));
            _userControl = view as UserControl;
            if (settingsProvider != null)
            {
                _settings = settingsProvider.Create();
            }
            _window = CreateToolWindow(view);
        }

        public UserControl UserControl { get { return _userControl; } }

        private readonly IVBE _vbe;
        protected IVBE VBE { get { return _vbe; } }

        private object _userControlObject;

        private IWindow CreateToolWindow(IDockableUserControl control)
        {
            IWindow toolWindow;
            try
            {
                var info = _vbe.Windows.CreateToolWindow(_addin, _DockableWindowHost.RegisteredProgId, control.Caption, control.ClassId);
                _userControlObject = info.UserControl;
                toolWindow = info.ToolWindow;
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

        public virtual void Show()
        {
            _window.IsVisible = true;
        }

        public void Hide()
        {
            _window.IsVisible = false;
        }

        ~DockableToolwindowPresenter()
        {
            Debug.WriteLine("DockableToolwindowPresenter finalized.");
        }
    }
}
