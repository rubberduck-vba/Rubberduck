using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public abstract class DockablePresenterBase : IDisposable
    {
        private readonly AddIn _addin;
        private readonly Window _window;
        protected readonly UserControl UserControl;

        protected DockablePresenterBase(VBE vbe, AddIn addin, IDockableUserControl control)
        {
            _vbe = vbe;
            _addin = addin;
            UserControl = control as UserControl;
            _window = CreateToolWindow(control);
        }

        private readonly VBE _vbe;
        protected VBE VBE { get { return _vbe; } }

        private Window CreateToolWindow(IDockableUserControl control)
        {
            object userControlObject = null;
            var toolWindow = _vbe.Windows.CreateToolWindow(_addin, DockableWindowHost.RegisteredProgId, control.Caption, control.ClassId, ref userControlObject);
            
            var userControlHost = (DockableWindowHost)userControlObject;
            toolWindow.Visible = true; //window resizing doesn't work without this

            EnsureMinimumWindowSize(toolWindow);

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

        public virtual void Dispose()
        {
            UserControl.Dispose();
        }
    }
}