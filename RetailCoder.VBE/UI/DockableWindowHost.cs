using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.Common.WinAPI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.WindowsApi;
using User32 = Rubberduck.Common.WinAPI.User32;

namespace Rubberduck.UI
{
    [ComVisible(true)]
    [Guid(RubberduckGuid.DockableWindowHostGuid)]
    [ProgId(RubberduckProgId.DockableWindowHostProgId)]    
    [EditorBrowsable(EditorBrowsableState.Never)]
    //Nothing breaks because we declare a ProgId
    // ReSharper disable once InconsistentNaming
    //Underscores make classes invisible to VB6 object explorer
    public partial class _DockableWindowHost : UserControl
    {       
        public static string RegisteredProgId => RubberduckProgId.DockableWindowHostProgId;

        // ReSharper disable UnusedAutoPropertyAccessor.Local
        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {            
            public int Left { get; set; }           
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }
        }
        // ReSharper restore UnusedAutoPropertyAccessor.Local

        [StructLayout(LayoutKind.Explicit)]
        private struct LParam
        {
            [FieldOffset(0)]
            public uint Value;
            [FieldOffset(0)]
            public readonly ushort LowWord;
            [FieldOffset(2)]
            public readonly ushort HighWord;
        }

        [DllImport("User32.dll")]
        static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("User32.dll", EntryPoint = "GetClientRect")]
        static extern int GetClientRect(IntPtr hWnd, ref Rect lpRect);

        private IntPtr _parentHandle;
        private ParentWindow _subClassingWindow;
        private GCHandle _thisHandle;

        internal void AddUserControl(UserControl control, IntPtr vbeHwnd)
        {
            _parentHandle = GetParent(Handle);
            _subClassingWindow = new ParentWindow(vbeHwnd, new IntPtr(GetHashCode()), _parentHandle);
            _subClassingWindow.CallBackEvent += OnCallBackEvent;

            //DO NOT REMOVE THIS CALL. Dockable windows are instantiated by the VBE, not directly by RD.  On top of that,
            //since we have to inherit from UserControl we don't have to keep handling window messages until the VBE gets
            //around to destroying the control's host or it results in an access violation when the base class is disposed.
            //We need to manually call base.Dispose() ONLY in response to a WM_DESTROY message.
            _thisHandle = GCHandle.Alloc(this, GCHandleType.Normal);

            if (control != null)
            {
                control.Dock = DockStyle.Fill;
                Controls.Add(control);
            }
            AdjustSize();
        }

        private void OnCallBackEvent(object sender, SubClassingWindowEventArgs e)
        {
            if (!e.Closing)
            {
                var param = new LParam {Value = (uint) e.LParam};
                Size = new Size(param.LowWord, param.HighWord);
            }
            else
            {
                Debug.WriteLine("DockableWindowHost removed event handler.");
                _subClassingWindow.CallBackEvent -= OnCallBackEvent;
            }
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

        protected override void DefWndProc(ref Message m)
        {
            //See the comment in the ctor for why we have to listen for this.
            if (m.Msg == (int) WM.DESTROY)
            {
                Debug.WriteLine("DockableWindowHost received WM.DESTROY.");
                _thisHandle.Free();
            }
            base.DefWndProc(ref m);
        }

        //override 

        public void Release()
        {
            Debug.WriteLine("DockableWindowHost release called.");
            _subClassingWindow.Dispose();
        }

        protected override void DestroyHandle()
        {
            Debug.WriteLine("DockableWindowHost DestroyHandle called.");
            base.DestroyHandle();
        }

        [ComVisible(false)]
        public class ParentWindow : SubclassingWindow
        {
            public event SubClassingWindowEventHandler CallBackEvent;
            public delegate void SubClassingWindowEventHandler(object sender, SubClassingWindowEventArgs e);

            private readonly IntPtr _vbeHwnd;

            private void OnCallBackEvent(SubClassingWindowEventArgs e)
            {
                if (CallBackEvent != null)
                {
                    CallBackEvent(this, e);
                }
            }
            
            public ParentWindow(IntPtr vbeHwnd, IntPtr id, IntPtr handle) : base(id, handle)
            {
                _vbeHwnd = vbeHwnd;
            }

            private bool _closing;
            public override int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
            {
                switch ((uint)msg)
                {
                    case (uint)WM.SIZE:
                        var args = new SubClassingWindowEventArgs(lParam);
                        if (!_closing) OnCallBackEvent(args);
                        break;
                    case (uint)WM.SETFOCUS:
                        if (!_closing) User32.SendMessage(_vbeHwnd, WM.RUBBERDUCK_CHILD_FOCUS, Hwnd, Hwnd);
                        break;
                    case (uint)WM.KILLFOCUS:
                        if (!_closing) User32.SendMessage(_vbeHwnd, WM.RUBBERDUCK_CHILD_FOCUS, Hwnd, IntPtr.Zero);
                        break;
                    case (uint)WM.RUBBERDUCK_SINKING:
                        OnCallBackEvent(new SubClassingWindowEventArgs(lParam) { Closing = true });
                        _closing = true;
                        break;
                }
                return base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
            }
        }
    }
}
