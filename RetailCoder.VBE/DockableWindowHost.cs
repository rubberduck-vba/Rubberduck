using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Runtime.InteropServices;

namespace RetailCoderVBE
{
    //todo: store GUID in a const so if it needs to be changed it can be changed once??
    //todo: needs a way to handle toolwindow scrollbars when resizing
    [ComVisible(true), Guid("9CF1392A-2DC9-48A6-AC0B-E601A9802608"), ProgId("RetailCoderVBE.DockableWindowHost")]
    public partial class DockableWindowHost : UserControl
    {
        private class SubClassingWindow : System.Windows.Forms.NativeWindow
        {

            public event SubClassingWindowEventHandler CallBackEvent;
            public delegate void SubClassingWindowEventHandler(object sender, SubClassingWindowEventArgs e);
           
            protected virtual void OnCallBackEvent(SubClassingWindowEventArgs e)
            {
                CallBackEvent(this, e);
            }
            

            public SubClassingWindow(IntPtr handle)
            {
                base.AssignHandle(handle);
            }

            protected override void WndProc(ref Message msg)
            {
                const int WM_SIZE = 0x5;

                if (msg.Msg == WM_SIZE)
                {
                    SubClassingWindowEventArgs args = new SubClassingWindowEventArgs(msg);
                    OnCallBackEvent(args);
                }

 	             base.WndProc(ref msg);
            }

            //destructor
            ~SubClassingWindow()
            {
                this.ReleaseHandle();
            }

        }//end  subclassingwindow class

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            internal int Left;
            internal int Top;
            internal int Right;
            internal int Bottom;

        }

        [DllImport("User32.dll")]
        static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("User32.dll", EntryPoint = "GetClientRect")]
        static extern int GetClientRect(IntPtr hWnd, ref RECT lpRect);

        private IntPtr parentHandle;
        private SubClassingWindow subClassingWindow;

        internal void AddUserControl(UserControl control)
        {
            parentHandle = GetParent(this.Handle);
            subClassingWindow = new SubClassingWindow(parentHandle);
            subClassingWindow.CallBackEvent += OnCallBackEvent;

            control.Dock = DockStyle.Fill;
            this.Controls.Add(control);

            AdjustSize();

        }

        void OnCallBackEvent(object sender, SubClassingWindowEventArgs e)
        {
            AdjustSize();
        }

        private void AdjustSize()
        {
            RECT tRect = new RECT();
            if (GetClientRect(parentHandle,ref tRect) != 0)
            {
                this.Size = new Size(tRect.Right - tRect.Left, tRect.Bottom - tRect.Top);
            }
        }

        protected override bool ProcessKeyPreview(ref Message m)
        {
            const int WM_KEYDOWN = 0x100;
            bool result = false;
            Keys pressedKey;
            UserControl hostedUserControl;
            Button activeButton;

            hostedUserControl = (UserControl)this.Controls[0];

            if (m.Msg == WM_KEYDOWN)
            {
                pressedKey = (Keys)m.WParam;
                switch(pressedKey)
                {
                    case Keys.Tab:
                        if (Control.ModifierKeys == Keys.None) //just tab
                        {
                            this.SelectNextControl(hostedUserControl.ActiveControl, true, true, true, true);
                            result = true;
                        }
                        else if (Control.ModifierKeys == Keys.Shift) //shift + tab
                        {
                            this.SelectNextControl(hostedUserControl.ActiveControl, false, true, true, true);
                            result = true;
                        }
                        break;
                    case Keys.Return:
                        if (hostedUserControl.ActiveControl.GetType().Equals(typeof(Button)))
                        {
                            activeButton = (Button)hostedUserControl.ActiveControl;
                            activeButton.PerformClick();
                        }
                        break;
                }
            }

            if (!result)
            {
                result = base.ProcessKeyPreview(ref m);
            }
            return result;
        }
    }

    internal class SubClassingWindowEventArgs : EventArgs
    {
        private Message msg;

        public Message Message
        {
            get { return this.msg; }
        }

        public SubClassingWindowEventArgs(System.Windows.Forms.Message msg)
        {
            this.msg = msg;
        }

    }
}
