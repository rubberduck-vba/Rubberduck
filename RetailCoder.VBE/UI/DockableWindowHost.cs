using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.Common.WinAPI;
using Rubberduck.VBEditor;

namespace Rubberduck.UI
{
    [Guid(ClassId)]
    [ProgId(ProgId)]
    [ComVisible(true)]
    [EditorBrowsable(EditorBrowsableState.Never)]

    //Nothing breaks because we declare a ProgId
    // ReSharper disable once InconsistentNaming
    //Underscores make classes invisible to VB6 object explorer
    public partial class _DockableWindowHost : UserControl
    {
        private const string ClassId = "9CF1392A-2DC9-48A6-AC0B-E601A9802608";
        private const string ProgId = "Rubberduck.UI.DockableWindowHost";
        public static string RegisteredProgId { get { return ProgId; } }

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

        internal void AddUserControl(UserControl control, IntPtr vbeHwnd)
        {
            _parentHandle = GetParent(Handle);
            _subClassingWindow = new SubClassingWindow(_parentHandle, vbeHwnd);
            _subClassingWindow.CallBackEvent += OnCallBackEvent;

            if (control != null)
            {
                control.Dock = DockStyle.Fill;
                Controls.Add(control);
            }
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
                switch (pressedKey)
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

            private readonly IntPtr _vbeHwnd;

            private void OnCallBackEvent(SubClassingWindowEventArgs e)
            {
                Debug.Assert(CallBackEvent != null, "CallBackEvent != null");
                CallBackEvent(this, e);
            }
            
            public SubClassingWindow(IntPtr handle, IntPtr vbeHwnd)
            {
                _vbeHwnd = vbeHwnd;
                AssignHandle(handle);
            }

            protected override void WndProc(ref Message msg)
            {
                var closing = false;
                switch ((uint)msg.Msg)
                {
                    case (uint)WM.SIZE:
                        var args = new SubClassingWindowEventArgs(msg);
                        OnCallBackEvent(args);
                        break;
                    case (uint)WM.SETFOCUS:
                        User32.SendMessage(_vbeHwnd, WM.RUBBERDUCK_CHILD_FOCUS, Handle, Handle);
                        break;
                    case (uint)WM.KILLFOCUS:
                        User32.SendMessage(_vbeHwnd, WM.RUBBERDUCK_CHILD_FOCUS, Handle, IntPtr.Zero);
                        break;
                    case (uint)WM.RUBBERDUCK_SINKING:
                        closing = true;
                        break;
                }
                base.WndProc(ref msg);
                if (closing)
                {
                    DestroyHandle();
                }
            }

            //~SubClassingWindow()
            //{
            //    ReleaseHandle();
            //}
        }
    }
}
