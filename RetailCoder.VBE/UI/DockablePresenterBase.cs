using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI
{
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
            try
            {
                object userControlObject = null;
                var toolWindow = _vbe.Windows.CreateToolWindow(_addin, _DockableWindowHost.RegisteredProgId, control.Caption, control.ClassId, ref userControlObject);

                var userControlHost = (_DockableWindowHost)userControlObject;
                toolWindow.Visible = true; //window resizing doesn't work without this

                EnsureMinimumWindowSize(toolWindow);

                userControlHost.AddUserControl(control as UserControl);
                return toolWindow;
            }
            catch (Exception)
            {
                // bug: there's a COM exception here if the window was X-closed before. see issue #169.
                return null;
            }
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
            try
            {
                _window.Visible = true;
            }
            catch (COMException e)
            {
                // bug: this exception shouldn't be happening. see issue #169.
            }
            catch (NullReferenceException e)
            {
                // bug: this exception shouldn't be happening either. may be related to #169... or not.
            }
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