using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Rubberduck.UI
{
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    public partial class DockableWindowHost : UserControl
    {
        public const string ClassId = "9CF1392A-2DC9-48A6-AC0B-E601A9802608";
        public const string ProgId = "Rubberduck.UI.DockableWindowHost";

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left { get; set; }
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }
        }

        [DllImport("User32.dll")]
        static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("User32.dll", EntryPoint = "GetClientRect")]
        static extern int GetClientRect(IntPtr hWnd, ref Rect lpRect);

        private IntPtr _parentHandle;
        private SubClassingWindow _subClassingWindow;

        internal void AddUserControl(UserControl control)
        {
            _parentHandle = GetParent(Handle);
            _subClassingWindow = new SubClassingWindow(_parentHandle);
            _subClassingWindow.CallBackEvent += OnCallBackEvent;

            control.Dock = DockStyle.Fill;
            Controls.Add(control);

            AdjustSize();
        }

        private void OnCallBackEvent(object sender, SubClassingWindowEventArgs e)
        {
            AdjustSize();
        }

        private void AdjustSize()
        {
            var rect = new Rect();
            if (GetClientRect(_parentHandle, ref rect) != 0)
            {
                Size = new Size(rect.Right - rect.Left, rect.Bottom - rect.Top);
            }
        }

        protected override bool ProcessKeyPreview(ref Message m)
        {
            const int wmKeydown = 0x100;
            var result = false;

            var hostedUserControl = (UserControl)Controls[0];

            if (m.Msg == wmKeydown)
            {
                var pressedKey = (Keys)m.WParam;
                switch(pressedKey)
                {
                    case Keys.Tab:
                        switch (ModifierKeys)
                        {
                            case Keys.None:
                                SelectNextControl(hostedUserControl.ActiveControl, true, true, true, true);
                                result = true;
                                break;
                            case Keys.Shift:
                                SelectNextControl(hostedUserControl.ActiveControl, false, true, true, true);
                                result = true;
                                break;
                        }
                        break;
                    case Keys.Return:
                        if (hostedUserControl.ActiveControl.GetType() == typeof(Button))
                        {
                            var activeButton = (Button)hostedUserControl.ActiveControl;
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

        [ComVisible(false)]
        public class SubClassingWindow : NativeWindow
        {
            public event SubClassingWindowEventHandler CallBackEvent;
            public delegate void SubClassingWindowEventHandler(object sender, SubClassingWindowEventArgs e);

            private void OnCallBackEvent(SubClassingWindowEventArgs e)
            {
                CallBackEvent(this, e);
            }
            
            public SubClassingWindow(IntPtr handle)
            {
                AssignHandle(handle);
            }

            protected override void WndProc(ref Message msg)
            {
                const int wmSize = 0x5;

                if (msg.Msg == wmSize)
                {
                    var args = new SubClassingWindowEventArgs(msg);
                    OnCallBackEvent(args);
                }

                base.WndProc(ref msg);
            }

            ~SubClassingWindow()
            {
                ReleaseHandle();
            }
        }
    }
}
