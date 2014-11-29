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

            userControlHost.AddUserControl(control as UserControl);
            return toolWindow;
        }

        public void Show()
        {
            _window.Visible = true;
        }

        public void Close()
        {
            _window.Close();
        }

        public void Dispose()
        {
            UserControl.Dispose();
        }
    }
}