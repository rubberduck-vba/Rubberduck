using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public abstract class DockablePresenterBase
    {
        public const string DockableWindowHostProgId = "Rubberduck.UI.DockableWindowHost";
        public const string DockableWindowHostClassId = "9CF1392A-2DC9-48A6-AC0B-E601A9802608";

        private readonly AddIn _addin;
        protected readonly Window Window;
        protected readonly UserControl UserControl;

        protected DockablePresenterBase(VBE vbe, AddIn addin, string caption, UserControl control)
        {
            _vbe = vbe;
            _addin = addin;
            UserControl = control;
            Window = CreateToolWindow(caption, control);
        }

        private readonly VBE _vbe;
        protected VBE VBE { get { return _vbe; } }

        private Window CreateToolWindow(string toolWindowCaption, UserControl toolWindowUserControl)
        {
            Object userControlObject = null;
            var toolWindow = _vbe.Windows.CreateToolWindow(_addin, DockableWindowHostProgId, toolWindowCaption, DockableWindowHostClassId, ref userControlObject);

            var userControlHost = (DockableWindowHost)userControlObject;
            toolWindow.Visible = true; //window resizing doesn't work without this

            userControlHost.AddUserControl(toolWindowUserControl);
            return toolWindow;
        }

        public void Show()
        {
            Window.Visible = true;
        }

        public void Close()
        {
            Window.Close();
        }
    }
}