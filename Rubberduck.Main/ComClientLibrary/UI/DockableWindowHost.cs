using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.WindowsApi;
using User32 = Rubberduck.Common.WinAPI.User32;
using NLog;
using Rubberduck.Resources.Registration;
using Rubberduck.UI.CustomComWrappers;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI
{
    // This is a new implementation of _DockabaleWindowHost, implementing the IOle* interfaces required for ActiveX support, and acts 
    // as a proxy between the host (VBE) and the UserControl.  By doing it this way, you can capture all the host-owned COM objects passed around
    // (through the ActiveX interfaces) and at shutdown you can release them all thoroughly and correctly.
    //
    // There are two main difficulties with this approach.  1) getting access to the internal IOle* interfaces implemented by the UserControl
    // so that we can forward calls on to them.  2) getting the CLR to accept our own versions of the IOle* interfaces in place of the internal ones 
    // defined in the System.Windows.Forms.UnsafeNativeMethods assembly.
    //
    // Both of these issues are resolved neatly using aggregation objects (Marshal.CreateAggregationObject), but we have to largely use raw
    // pointers (IntPtr) to pass around the COM interfaces, and so great care has to be taken with ensuring explicit reference counting is 
    // COM compliant.
    //
    // CLASSES:
    // _DockableWindowHost: the main class, implementing the proxy IOleObject and IOleInPlaceObject interfaces
    // ExposedUserControl: exposes the internal IOleObject and IOleInPlaceObject interfaces from System.Windows.Forms.UnsafeNativeMethods
    // the rest are wrapper classes and helpers.
    // The old subclassing code is still used, at the bottom of the new _DockableWindowHost class.
    //
    // All defined IOle* interfaces are _compatible_ [not identical] to the originals.  Aggregation objects are used to pass interfaces 
    // from our own versions of the IOle* interfaces to the UserControl ones (defined in System.Windows.Forms.UnsafeNativeMethods)
    //  In particular, our interfaces differ by defining all pointers, including object instances, as IntPtr
    //   this is for two reasons: 
    //    1) we need to support nullified reference input params.  C# does not support them, and the CLR will throw exceptions with them.
    //    2) we have to avoid the CLR resolving the aggregated objects to our own versions of the interfaces whenever possible
    //   - this means that manual reference counting (with Marshal.AddRef / Release) is in place to ensure proper COM compliance
    //
    // Differences from the old DockableWindowHost:
    //   we now have a proper IOleObject::Close() event that is the last thing to fire on our object.  We should use this to tear down
    //   all the COM instances obtained for each instance of DockableWindowHost.  This is equivalent to the OnDisconnection event on the 
    //   main AddIn side, and we should respect that the VBE expects us to release everything relating to the ActiveX control at this exact point.
    //   IMO the main addin should NOT now trigger our destruction process.  It should be self contained, using the IOleObject::Close event.
    //
    // TODO improve comments
    // TODO some C# magic?  (I only dabble in C#, so I'm sure this isn't going to win awards for the best use of C# features ;)
    //
    // -- Wayne Phillips 29th Dec 2017
    //
    //
    // Note that the wrapper infrastructure has been moved to the namespace Rubberduck.VBEditor.CustomComWrappers.
    //
    //

    [
        ComVisible(true),
        Guid(RubberduckGuid.IDockableWindowHostGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Never)
    ]
    public interface IDockableWindowHost
    {
        [DispId(1)]
        void AddUserControl(UserControl control, IntPtr vbeHwnd);
        [DispId(2)]
        void Release();
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.DockableWindowHostGuid),
        ProgId(RubberduckProgId.DockableWindowHostProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IDockableWindowHost)),
        EditorBrowsable(EditorBrowsableState.Never)
    ]
    //Nothing breaks because we declare a ProgId
    // ReSharper disable once InconsistentNaming
    //Underscores make classes invisible to VB6 object explorer
    [SuppressMessage("Microsoft.Design", "CA1049")]
    [SuppressMessage("Microsoft.Design", "CA1001")] //This should *never* have Dispose called on it. See comment block above.
    public class _DockableWindowHost : COM_IOleObject, COM_IOleInPlaceObject, COM_IOleWindow, IDockableWindowHost
    {
        public static string RegisteredProgId => RubberduckProgId.DockableWindowHostProgId;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private ExposedUserControl _userControl = new ExposedUserControl();
        private Wrapper_IOleClientSite _cachedClientSite;

        private bool _releaseHasBeenCalled;
        public void Release()
        {
            // WARNING: Disposal of _userControl / _cachedClientSite should be handled in IOleObject::Close(), not here, see top comments
            _releaseHasBeenCalled = true;
            RemoveChildControlsFromExposedControl();
        }

        private void RemoveChildControlsFromExposedControl()
        {
            while (_userControl.Controls.Count > 0)
            {
                var control = _userControl.Controls[0];
                _userControl.Controls.Remove(control);
                control.Dispose();
            }
        }


        public int /* IOleObject:: */ SetClientSite([In] IntPtr /* IOleClientSite */ pClientSite)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::SetClientSite() called");

            // IOleObject::SetClientSite is typically called with pClientSite = null just before calling IOleObject::Close()
            // We release all host COM objects here in that instance.

            _cachedClientSite?.Dispose();
            _cachedClientSite = null;

            if (pClientSite != IntPtr.Zero)
            {
                _cachedClientSite = new Wrapper_IOleClientSite(pClientSite);
                return _userControl.IOleObject.SetClientSite(_cachedClientSite.PeekAggregatedReference());     // callee will take its own reference
            }
            return (int)ComConstants.S_OK;
        }

        public int /* IOleObject:: */ GetClientSite([Out] out IntPtr /* IOleClientSite */ ppClientSite)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetClientSite() called");
            ppClientSite = _cachedClientSite?.CopyAggregatedReference() ?? IntPtr.Zero;
            return (int)ComConstants.S_OK;
        }

        public int /* IOleObject:: */ SetHostNames([In, MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [In, MarshalAs(UnmanagedType.LPWStr)] string szContainerObj)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::SetHostNames() called");
            return _userControl.IOleObject.SetHostNames(szContainerApp, szContainerObj);
        }

        public int /* IOleObject:: */ Close([In] uint dwSaveOption)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Close() called");
            var hr = _userControl.IOleObject.Close(dwSaveOption);

            // IOleObject::SetClientSite is typically called with pClientSite = null just before calling IOleObject::Close()
            // If it didn't, we release all host COM objects here instead,

            // This is the point where we can deterministically, and safely release our COM references for this ActiveX control.
            // Moreover, we can release the UserControl COM references, as Close() should be the very last call into the IOleObject interface.
            PerformUserControlShutdown();

            return hr;
        }

        private void PerformUserControlShutdown()
        {
            ReleaseActiveXControlComReference();
            ReleasedExposedControl();
            UnsubclassParent();
        }

        private void UnsubclassParent()
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Close() ... unsubclassing the host parent window");
            _subClassingWindow.CallBackEvent -= OnCallBackEvent;
            _subClassingWindow.Dispose();
        }

        private void ReleasedExposedControl()
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Close() ... closing down internal COM references");
            _userControl.Dispose();
            _userControl = null;
        }

        private void ReleaseActiveXControlComReference()
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Close() ... closing down host COM references");
            _cachedClientSite?.Dispose();
            _cachedClientSite = null;
        }

        public int /* IOleObject:: */ SetMoniker([In] uint dwWhichMoniker, [In] IntPtr /* IMoniker */ pmk)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::SetMoniker() called");
            // need to wrap IMoniker to support this.  Not used by VBE anyway?
            //return _IOleObject.SetMoniker(dwWhichMoniker, pmk);
            Debug.Assert(false);
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ GetMoniker([In] uint dwAssign, [In] uint dwWhichMoniker, [Out] out IntPtr /* IMoniker */ ppmk)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetMoniker() called");
            // need to wrap IMoniker to support this.  Not used by VBE anyway?
            //return _IOleObject.GetMoniker(dwAssign, dwWhichMoniker, out ppmk);
            ppmk = IntPtr.Zero;
            Debug.Assert(false);
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ InitFromData([In] IntPtr /* IDataObject */ pDataObject, [In] int fCreation, [In] uint dwReserved)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::InitFromData() called");
            // need to wrap IDataObject to support this.  Not used by VBE anyway?
            //return _IOleObject.InitFromData(pDataObject, fCreation, dwReserved);
            Debug.Assert(false);
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ GetClipboardData([In] uint dwReserved, [Out] out IntPtr /*IDataObject*/ ppDataObject)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetClipboardData() called");
            // need to wrap IDataObject to support this.  Not used by VBE anyway?
            //return _IOleObject.GetClipboardData(dwReserved, out ppDataObject);
            ppDataObject = IntPtr.Zero;
            Debug.Assert(false);
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ DoVerb([In] int iVerb, [In] IntPtr lpmsg, [In] IntPtr /* IOleClientSite */ pActiveSite, [In] int lindex, [In] IntPtr hwndParent, [In] IntPtr /* COMRECT */ lprcPosRect)
        {
            _logger.Log(LogLevel.Trace, $"IOleObject::DoVerb() called with iVerb {(Enum.IsDefined(typeof(OleVerbs), iVerb) ? ((OleVerbs) iVerb).ToString() : iVerb.ToString())}.");
            // pActiveSite is not used by the UserControl implementation.  Either wrap it or pass null instead
            pActiveSite = IntPtr.Zero;

            //note: We swallow this OleVerb after release has been called because it is causing problems on shutdown.
            if (_releaseHasBeenCalled && iVerb == (int)OleVerbs.OLEIVERB_DISCARDUNDOSTATE)
            {
                return (int)ComConstants.S_OK;
            }

            return _userControl.IOleObject.DoVerb(iVerb, lpmsg, pActiveSite, lindex, hwndParent, lprcPosRect);
        }

        public int /* IOleObject:: */ EnumVerbs([Out] out IntPtr /* IEnumOLEVERB */ ppEnumOleVerb)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::EnumVerbs() called");
            // need to wrap IEnumOLEVERB to support this.  Not used by VBE anyway?
            //return _IOleObject.EnumVerbs(out ppEnumOleVerb);
            ppEnumOleVerb = IntPtr.Zero;
            Debug.Assert(false);
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ Update()
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Update() called");
            return _userControl.IOleObject.Update();
        }

        public int /* IOleObject:: */ IsUpToDate()
        {
            _logger.Log(LogLevel.Trace, "IOleObject::IsUpToDate() called");
            return _userControl.IOleObject.IsUpToDate();
        }

        public int /* IOleObject:: */ GetUserClassID([Out] out Guid pClsid)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetUserClassID() called");
            return _userControl.IOleObject.GetUserClassID(out pClsid);
        }

        public int /* IOleObject:: */ GetUserType([In] uint dwFormOfType, [Out, MarshalAs(UnmanagedType.LPWStr)] out string pszUserType)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetUserType() called");
            return _userControl.IOleObject.GetUserType(dwFormOfType, out pszUserType);
        }

        public int /* IOleObject:: */ SetExtent([In] uint dwDrawAspect, [In] IntPtr /* tagSIZE */ psizel)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::SetExtent() called");
            return _userControl.IOleObject.SetExtent(dwDrawAspect, psizel);
        }

        public int /* IOleObject:: */ GetExtent([In] uint dwDrawAspect, [Out] out IntPtr /* tagSIZE */ psizel)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetExtent() called");
            return _userControl.IOleObject.GetExtent(dwDrawAspect, out psizel);
        }

        public int /* IOleObject:: */ Advise([In] IntPtr /* IAdviseSink */ pAdvSink, [Out] out uint pdwConnection)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Advise() called");
            // need to wrap IAdviseSink to support this. VBE does try to use this, but the events don't look interesting?
            pdwConnection = 0;
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ Unadvise([In] uint pdwConnection)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Unadvise() called");
            // No sense supporting Unadvise, as we're not supporting Advise
            //return _IOleObject.Unadvise(pdwConnection);
            //Debug.Assert(false);                              stupid VBE still calls us, despite us not implementing Advise()
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ EnumAdvise([Out] out IntPtr /* IEnumSTATDATA */ enumAdvise)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::EnumAdvise() called");
            // need to wrap IEnumSTATDATA to support this. No sense supporting EnumAdvise, as we're not supporting Advise
            //return _IOleObject.EnumAdvise(out enumAdvise);
            enumAdvise = IntPtr.Zero;
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ GetMiscStatus([In] uint dwAspect, [Out] out uint pdwStatus)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetMiscStatus() called");
            return _userControl.IOleObject.GetMiscStatus(dwAspect, out pdwStatus);
        }

        public int /* IOleObject:: */ SetColorScheme([In] IntPtr /* tagLOGPALETTE */ pLogpal)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::SetColorScheme() called");
            return _userControl.IOleObject.SetColorScheme(pLogpal);
        }

        // --------------------------------------------------------------------

        public int /* IOleInPlaceObject:: */ GetWindow([Out] out IntPtr hwnd)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::GetWindow() called");
            return _userControl.IOleInPlaceObject.GetWindow(out hwnd);
        }

        public int /* IOleInPlaceObject:: */ ContextSensitiveHelp([In] int fEnterMode)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::ContextSensitiveHelp() called");
            return _userControl.IOleInPlaceObject.ContextSensitiveHelp(fEnterMode);
        }

        public int /* IOleInPlaceObject:: */ InPlaceDeactivate()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::InPlaceDeactivate() called");
            return _userControl.IOleInPlaceObject.InPlaceDeactivate();
        }

        public int /* IOleInPlaceObject:: */ UIDeactivate()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::UIDeactivate() called");
            return _userControl.IOleInPlaceObject.UIDeactivate();
        }

        public int /* IOleInPlaceObject:: */ SetObjectRects([In] IntPtr lprcPosRect, [In] IntPtr lprcClipRect)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::SetObjectRects() called");
            return _userControl.IOleInPlaceObject.SetObjectRects(lprcPosRect, lprcClipRect);
        }

        public int /* IOleInPlaceObject:: */ ReactivateAndUndo()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::ReactivateAndUndo() called");
            return _userControl.IOleInPlaceObject.ReactivateAndUndo();
        }

        // old stuff from old _DockableWindowHost --------------------------------------- [START]

        private void OnCallBackEvent(object sender, SubClassingWindowEventArgs e)
        {
            if (e.Closing)
            {
                return;
            }
            var param = new LParam { Value = (uint)e.LParam };
            // The VBE passes a special value to the HighWord when docking into the VBE Codepane 
            // instead of docking into the VBE itself.
            // that special value (0xffef) blows up inside the guts of Window management because it's
            // apparently converted to a signed short somewhere and then considered "negative"
            // that is why we drop the signbit for shorts off our values when creating the control Size.
            const ushort signBitMask = 0x8000;
            _userControl.Size = new Size(param.LowWord & ~signBitMask, param.HighWord & ~signBitMask);
        }

        public void AddUserControl(UserControl control, IntPtr vbeHwnd)
        {
            _parentHandle = GetParent(_userControl.Handle);
            _subClassingWindow = new ParentWindow(vbeHwnd, new IntPtr(GetHashCode()), _parentHandle);
            _subClassingWindow.CallBackEvent += OnCallBackEvent;

            control.Dock = DockStyle.Fill;
            _userControl.Controls.Add(control);
            
            AdjustSize();
        }

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

        private void AdjustSize()
        {
            var rect = new Rect();
            if (GetClientRect(_parentHandle, ref rect) != 0)
            {
                _userControl.Size = new Size(rect.Right - rect.Left, rect.Bottom - rect.Top);
            }
        }

        private static void ToggleDockable(IntPtr hWndVBE)
        {
            NativeMethods.SendMessage(hWndVBE, 0x1044, (IntPtr)0xB5, IntPtr.Zero);
        }

        [ComVisible(false)]
        public class ParentWindow : SubclassingWindow
        {
            private readonly Logger _logger = LogManager.GetCurrentClassLogger();

            private const int MF_BYPOSITION = 0x400;

            public event SubClassingWindowEventHandler CallBackEvent;
            public delegate void SubClassingWindowEventHandler(object sender, SubClassingWindowEventArgs e);

            private readonly IntPtr _vbeHwnd;

            private IntPtr _containerHwnd;
            private ToolWindowState _windowState;
            private IntPtr _menuHandle;

            private enum ToolWindowState
            {
                Unknown,
                Docked,
                Floating,
                Undockable
            }

            private ToolWindowState GetWindowState(IntPtr containerHwnd)
            {
                var className = new StringBuilder(255);
                if (NativeMethods.GetClassName(containerHwnd, className, className.Capacity) > 0)
                {
                    switch (className.ToString())
                    {
                        case "wndclass_desked_gsk":
                            return ToolWindowState.Docked;
                        case "VBFloatingPalette":
                            return ToolWindowState.Floating;
                        case "DockingView":
                            return ToolWindowState.Undockable;
                    }
                }

                return ToolWindowState.Unknown;
            }

            private void DisplayUndockableContextMenu(IntPtr handle, IntPtr lParam)
            {
                if (_menuHandle == IntPtr.Zero)
                {
                    _menuHandle = NativeMethods.CreatePopupMenu();

                    if (_menuHandle == IntPtr.Zero)
                    {
                        _logger.Warn("Cannot create menu handle");
                        return;
                    }

                    if (!NativeMethods.InsertMenu(_menuHandle, 0, MF_BYPOSITION, (UIntPtr)WM.RUBBERDUCK_UNDOCKABLE_CONTEXT_MENU, "Dockable" + char.MinValue))
                    {
                        _logger.Warn("Failed to insert a menu item for dockable command");
                    }
                }

                var param = new LParam {Value = (uint)lParam};
                if (!NativeMethods.TrackPopupMenuEx(_menuHandle, 0x0, param.LowWord, param.HighWord, handle, IntPtr.Zero ))
                {
                    _logger.Warn("Failed to set the context menu for undockable tool windows");
                };
            }

            private void OnCallBackEvent(SubClassingWindowEventArgs e)
            {
                CallBackEvent?.Invoke(this, e);
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
                    case (uint)WM.WINDOWPOSCHANGED:
                        var containerHwnd = GetParent(hWnd);
                        if (containerHwnd != _containerHwnd)
                        {
                            _containerHwnd = containerHwnd;
                            _windowState = GetWindowState(_containerHwnd);
                        }
                        break;
                    case (uint)WM.CONTEXTMENU:
                        if (_windowState == ToolWindowState.Undockable)
                        {
                            DisplayUndockableContextMenu(hWnd, lParam);
                        }
                        break;
                    case (uint)WM.COMMAND:
                        switch (wParam.ToInt32())
                        {
                            case (int)WM.RUBBERDUCK_UNDOCKABLE_CONTEXT_MENU:
                                ToggleDockable(_vbeHwnd);
                                break;
                        }
                        break;
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
                    case (uint)WM.DESTROY:
                        if (_menuHandle != IntPtr.Zero)
                        {
                            if (!NativeMethods.DestroyMenu(_menuHandle))
                            {
                                _logger.Fatal($"Failed to destroy the menu handle {_menuHandle}");
                            }
                        }
                        break;
                }
                return base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
            }

            private bool _disposed;
            protected override void Dispose(bool disposing)
            {
                if (!_disposed && disposing && !_closing)
                {
                    OnCallBackEvent(new SubClassingWindowEventArgs(IntPtr.Zero) { Closing = true });
                    _closing = true;
                }

                _disposed = true;

                base.Dispose(disposing);
            }
        }

        // old stuff from old _DockableWindowHost --------------------------------------- [END]

        // ExposedUserControl - wrapper for UserControl that also exposes the underlying 
        // IOleObject and IOleInPlaceObject COM interfaces implemented by it
        public class ExposedUserControl : UserControl
        {
            private readonly Logger _logger = LogManager.GetCurrentClassLogger();

            public COM_IOleObject IOleObject;                  // cached interface obtained from UserConrol
            public COM_IOleInPlaceObject IOleInPlaceObject;    // cached interface obtained from UserConrol

            public ExposedUserControl()
            {
                _logger.Log(LogLevel.Trace, "ExposedUserControl constructor called");

                // Gain access to the IOleObject and IOleInPlaceObject interfaces implemented by the UserControl
                IOleObject = (COM_IOleObject)AggregationHelper.ObtainInternalInterface(this, GetType().GetInterface("IOleObject"));
                IOleInPlaceObject = (COM_IOleInPlaceObject)AggregationHelper.ObtainInternalInterface(this, GetType().GetInterface("IOleInPlaceObject"));
            }

            protected override void Dispose(bool disposing)
            {
                if (IOleObject != null)
                {
                    Marshal.ReleaseComObject(IOleObject);
                    IOleObject = null;
                }

                if (IOleInPlaceObject != null)
                {
                    Marshal.ReleaseComObject(IOleInPlaceObject);
                    IOleInPlaceObject = null;
                }

                base.Dispose(disposing);
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
        }
    }
}
