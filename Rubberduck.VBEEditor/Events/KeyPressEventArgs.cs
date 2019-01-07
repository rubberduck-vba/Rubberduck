using System;
using System.Windows.Forms;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.VBEditor.Events
{
    public class KeyPressEventArgs : EventArgs
    {
        public KeyPressEventArgs(IntPtr hwnd, IntPtr wParam, IntPtr lParam, bool keydown = false)
        {
            Hwnd = hwnd;
            WParam = wParam;
            LParam = lParam;
            ControlDown = (User32.GetKeyState(VirtualKeyStates.VK_CONTROL) & 0x8000) != 0;

            if (keydown)
            {
                if (((Keys) wParam & Keys.KeyCode) == Keys.Enter)
                {
                    // Why \r and not \n? Because it really doesn't matter...
                    Character = '\r';
                }
                else if (((Keys) wParam & Keys.KeyCode) == Keys.Back)
                {
                    Character = '\b';
                }
                else
                {
                    Character = default;
                }
            }
            else
            {              
                Character = (char)wParam;
            }
        }

        public IntPtr Hwnd { get; }
        public IntPtr WParam { get; }
        public IntPtr LParam { get; }

        public bool Handled { get; set; }
        public bool IsDelete => (Keys)WParam == Keys.Delete;
        public char Character { get; }
        public bool ControlDown { get; }
    }
}
