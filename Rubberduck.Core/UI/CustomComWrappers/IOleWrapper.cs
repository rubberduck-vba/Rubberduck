using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using NLog;

namespace Rubberduck.UI.CustomComWrappers
{
    // Relevant extract from the original comment by Wayne Phillips on 29th Dec 2017 from the DockableWindowHost before this got extracted here: 

    // ...
    //
    // There are two main difficulties with this approach.  1) getting access to the internal IOle* interfaces implemented by the UserControl
    // so that we can forward calls on to them.  2) getting the CLR to accept our own versions of the IOle* interfaces in place of the internal ones 
    // defined in the System.Windows.Forms.UnsafeNativeMethods assembly.
    //
    // Both of these issues are resolved neatly using aggregation objects (Marshal.CreateAggregationObject), but we have to largely use raw
    // pointers (IntPtr) to pass around the COM interfaces, and so great care has to be taken with ensuring explicit reference counting is 
    // COM compliant.
    //
    // ...
    //
    // All defined IOle* interfaces are _compatible_ [not identical] to the originals.  Aggregation objects are used to pass interfaces 
    // from our own versions of the IOle* interfaces to the UserControl ones (defined in System.Windows.Forms.UnsafeNativeMethods)
    //  In particular, our interfaces differ by defining all pointers, including object instances, as IntPtr
    //   this is for two reasons: 
    //    1) we need to support nullified reference input params.  C# does not support them, and the CLR will throw exceptions with them.
    //    2) we have to avoid the CLR resolving the aggregated objects to our own versions of the interfaces whenever possible
    //   - this means that manual reference counting (with Marshal.AddRef / Release) is in place to ensure proper COM compliance
    //
    // ...
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

    public enum ComConstants
    {
        E_NOTIMPL = -2147467263,
        S_OK = 0
    }

    public enum OleVerbs
    {
        OLEIVERB_SHOW = -1,
        OLEIVERB_OPEN = -2,
        OLEIVERB_HIDE = -3,
        OLEIVERB_UIACTIVATE = -4,
        OLEIVERB_INPLACEACTIVATE = -5,
        OLEIVERB_DISCARDUNDOSTATE = -6
    }

    [ComImport(), Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleObject
    {
        [PreserveSig] int SetClientSite([In] IntPtr /* IOleClientSite */ pClientSite);
        [PreserveSig] int GetClientSite([Out] out IntPtr /* IOleClientSite */ ppClientSite);
        [PreserveSig] int SetHostNames([In, MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [In, MarshalAs(UnmanagedType.LPWStr)] string szContainerObj);
        [PreserveSig] int Close([In] uint dwSaveOption);
        [PreserveSig] int SetMoniker([In] uint dwWhichMoniker, [In] IntPtr /* IMoniker */ pmk);
        [PreserveSig] int GetMoniker([In] uint dwAssign, [In] uint dwWhichMoniker, [Out] out IntPtr /* IMoniker */ ppmk);
        [PreserveSig] int InitFromData([In] IntPtr /* IDataObject */ pDataObject, [In] int fCreation, [In] uint dwReserved);
        [PreserveSig] int GetClipboardData([In] uint dwReserved, [Out] out IntPtr /*IDataObject*/ ppDataObject);
        [PreserveSig] int DoVerb([In] int iVerb, [In] IntPtr /* MSG, nullable ref */ lpmsg, [In] IntPtr /* IOleClientSite */ pActiveSite, [In] int lindex, [In] IntPtr hwndParent, [In] IntPtr /* COMRECT */ lprcPosRect);
        [PreserveSig] int EnumVerbs([Out] out IntPtr /* IEnumOLEVERB */ ppEnumOleVerb);
        [PreserveSig] int Update();
        [PreserveSig] int IsUpToDate();
        [PreserveSig] int GetUserClassID([Out] out Guid pClsid);
        [PreserveSig] int GetUserType([In] uint dwFormOfType, [Out, MarshalAs(UnmanagedType.LPWStr)] out string pszUserType);
        [PreserveSig] int SetExtent([In] uint dwDrawAspect, [In] IntPtr /* tagSIZE */ psizel);
        [PreserveSig] int GetExtent([In] uint dwDrawAspect, [Out] out IntPtr /* tagSIZE */ psizel);
        [PreserveSig] int Advise([In] IntPtr /* IAdviseSink */ pAdvSink, [Out] out uint pdwConnection);
        [PreserveSig] int Unadvise([In] uint pdwConnection);
        [PreserveSig] int EnumAdvise([Out] out IntPtr /* IEnumSTATDATA */ enumAdvise);
        [PreserveSig] int GetMiscStatus([In] uint dwAspect, [Out] out uint pdwStatus);
        [PreserveSig] int SetColorScheme([In] IntPtr /* tagLOGPALETTE */ pLogpal);
    };

    [ComImport(), Guid("00000113-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleInPlaceObject /* : COM_IOleWindow */
    {
        [PreserveSig] int GetWindow([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp([In] int fEnterMode);
        [PreserveSig] int InPlaceDeactivate();
        [PreserveSig] int UIDeactivate();
        [PreserveSig] int SetObjectRects([In] IntPtr /* COMRECT */ lprcPosRect, [In] IntPtr /* COMRECT */ lprcClipRect);
        [PreserveSig] int ReactivateAndUndo();
    }

    [ComImport(), Guid("00000118-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleClientSite
    {
        [PreserveSig] int SaveObject();
        [PreserveSig] int GetMoniker([In] uint dwAssign, [In] uint dwWhichMoniker, [Out] out IntPtr /* IMoniker */ moniker);
        [PreserveSig] int GetContainer([Out] out IntPtr /* IOleContainer */ container);
        [PreserveSig] int ShowObject();
        [PreserveSig] int OnShowWindow([In] int fShow);
        [PreserveSig] int RequestNewObjectLayout();
    }

    [ComImport(), Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleWindow
    {
        [PreserveSig] int GetWindow([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp([In] int fEnterMode);
    }

    [ComImport(), Guid("00000115-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleInPlaceUIWindow /* : COM_IOleWindow */
    {
        [PreserveSig] int GetWindow([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp([In] int fEnterMode);
        [PreserveSig] int GetBorder([Out] out IntPtr /* COMRECT */ lprectBorder);
        [PreserveSig] int RequestBorderSpace([In] IntPtr /* COMRECT */ pborderwidths);
        [PreserveSig] int SetBorderSpace([In] IntPtr /* COMRECT */ pborderwidths);
        [PreserveSig] int SetActiveObject([In] IntPtr /* IOleInPlaceActiveObject */ pActiveObject, [In, MarshalAs(UnmanagedType.LPWStr)] string pszObjName);
    }

    [ComImport(), Guid("00000116-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleInPlaceFrame /* : COM_IOleInPlaceUIWindow */
    {
        [PreserveSig] int GetWindow([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp([In] int fEnterMode);
        [PreserveSig] int GetBorder([Out] out IntPtr /* COMRECT */ lprectBorder);
        [PreserveSig] int RequestBorderSpace([In] IntPtr /* COMRECT */ pborderwidths);
        [PreserveSig] int SetBorderSpace([In] IntPtr /* COMRECT */ pborderwidths);
        [PreserveSig] int SetActiveObject([In] IntPtr /* IOleInPlaceActiveObject */ pActiveObject, [In, MarshalAs(UnmanagedType.LPWStr)] string pszObjName);
        [PreserveSig] int InsertMenus([In] IntPtr hmenuShared, [In, Out] ref IntPtr /* tagOleMenuGroupWidths */ lpMenuWidths);
        [PreserveSig] int SetMenu([In] IntPtr hmenuShared, [In] IntPtr holemenu, [In] IntPtr hwndActiveObject);
        [PreserveSig] int RemoveMenus([In] IntPtr hmenuShared);
        [PreserveSig] int SetStatusText([In, MarshalAs(UnmanagedType.LPWStr)] string pszStatusText);
        [PreserveSig] int EnableModeless([In] bool fEnable);
        [PreserveSig] int TranslateAccelerator([In] IntPtr lpmsg, [In] ushort wID);
    }

    [ComImport(), Guid("00000119-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleInPlaceSite /* : COM_IOleWindow */
    {
        [PreserveSig] int GetWindow([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp([In] int fEnterMode);
        [PreserveSig] int CanInPlaceActivate();
        [PreserveSig] int OnInPlaceActivate();
        [PreserveSig] int OnUIActivate();
        [PreserveSig] int GetWindowContext([Out] out IntPtr /* IOleInPlaceFrame */ ppFrame, [Out] out IntPtr /* IOleInPlaceUIWindow */ ppDoc, [Out] out IntPtr /* COMRECT */ lprcPosRect, [Out] out IntPtr /* COMRECT */ lprcClipRect, [In] IntPtr /* tagOIFI */ lpFrameInfo);
        [PreserveSig] int Scroll([In] IntPtr /* tagSIZE */ scrollExtant);
        [PreserveSig] int OnUIDeactivate([In] int fUndoable);
        [PreserveSig] int OnInPlaceDeactivate();
        [PreserveSig] int DiscardUndoState();
        [PreserveSig] int DeactivateAndUndo();
        [PreserveSig] int OnPosRectChange([In] IntPtr /* COMRECT */ lprcPosRect);
    }

    public class AggregationHelper : ICustomQueryInterface, IDisposable
    {
        private object _outerObject;
        private Type[] _supportedTypes;

        // CreateAggregatedWrapper returns a reference counted COM pointer to the aggregated object
        // When it gets released in COM, it should in turn release the internal CCW reference on our AggregationHelper object
        public static IntPtr CreateAggregatedWrapper(object objectToWrap, Type[] supportedTypes)
        {
            return Marshal.CreateAggregatedObject(Marshal.GetIUnknownForObject(objectToWrap),       // aggregated object will own this COM reference, but this is fine, as it is really a managed object
                new AggregationHelper(objectToWrap, supportedTypes));
        }

        // ObtainInternalInterface is used to obtain an object representing the internal interface implemented by an object
        //  (e.g. the internal IOleObject interface implemented by UserControl)
        // just cast the returned object to a local equivalent interface
        public static object ObtainInternalInterface(object objectInstance, Type internalType)
        {
            var aggregatedObjectPtr = CreateAggregatedWrapper(objectInstance, new Type[] { internalType });
            var retVal = Marshal.GetObjectForIUnknown(aggregatedObjectPtr);
            Marshal.Release(aggregatedObjectPtr);       // retVal holds an RCW reference now, so this can be released
            return retVal;
        }

        public AggregationHelper(object outerObject, Type[] supportedTypes)
        {
            _outerObject = outerObject;
            _supportedTypes = supportedTypes;
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            if (_outerObject != null)
            {
                foreach (Type _interface in _supportedTypes)
                {
                    if (_interface.GUID == iid)
                    {
                        ppv = Marshal.GetComInterfaceForObject(_outerObject, _interface, CustomQueryInterfaceMode.Ignore);
                        if (ppv != IntPtr.Zero)
                        {
                            return CustomQueryInterfaceResult.Handled;
                        }
                    }
                }
            }
            return CustomQueryInterfaceResult.Failed;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            _outerObject = null;
            _supportedTypes = null;

            _isDisposed = true;
        }
    }

    public class AggregatedWrapper : IDisposable
    {
        private IntPtr _aggregatedObjectPtr;

        // no explicit interface list defined, so use our implemented interfaces list
        public AggregatedWrapper()
        {
            _aggregatedObjectPtr = AggregationHelper.CreateAggregatedWrapper(this, GetType().GetInterfaces());
        }

        // interface list explicitly defined
        public AggregatedWrapper(Type[] supportedTypes)
        {
            _aggregatedObjectPtr = AggregationHelper.CreateAggregatedWrapper(this, supportedTypes);
        }

        public IntPtr CopyAggregatedReference()
        {
            Marshal.AddRef(_aggregatedObjectPtr);
            return _aggregatedObjectPtr;
        }

        public IntPtr PeekAggregatedReference()
        {
            return _aggregatedObjectPtr;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            if (_aggregatedObjectPtr != IntPtr.Zero)
            {
                Marshal.Release(_aggregatedObjectPtr);
                _aggregatedObjectPtr = IntPtr.Zero;
            }

            _isDisposed = true;
        }
    }

    public class WrapperBase : AggregatedWrapper
    {
        private IntPtr _hostObjectPtr;

        // no explicit interface list defined, so use our implemented interfaces list
        public WrapperBase(IntPtr hostObjectPtr)
        {
            _hostObjectPtr = hostObjectPtr;
            if (_hostObjectPtr != IntPtr.Zero) Marshal.AddRef(_hostObjectPtr);
        }

        // interface list explicitly defined
        public WrapperBase(IntPtr hostObjectPtr, Type[] supportedTypes) : base(supportedTypes)
        {
            _hostObjectPtr = hostObjectPtr;
            if (_hostObjectPtr != IntPtr.Zero) Marshal.AddRef(_hostObjectPtr);
        }

        public object GetObject()
        {
            return Marshal.GetObjectForIUnknown(_hostObjectPtr);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_hostObjectPtr != IntPtr.Zero)
                {
                    Marshal.Release(_hostObjectPtr);
                    _hostObjectPtr = IntPtr.Zero;
                }
            }

            base.Dispose(disposing);
        }
    }

    public class Wrapper_IOleInPlaceFrame : WrapperBase, COM_IOleInPlaceFrame, COM_IOleInPlaceUIWindow, COM_IOleWindow
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private COM_IOleInPlaceFrame _IOleInPlaceFrame;         // cached object for accessing the IOleInPlaceFrame interface

        public Wrapper_IOleInPlaceFrame(IntPtr hostObjectPtr) : base(hostObjectPtr)
        {
            if (hostObjectPtr != IntPtr.Zero)
            {
                _IOleInPlaceFrame = (COM_IOleInPlaceFrame)GetObject();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_IOleInPlaceFrame != null)
                {
                    Marshal.ReleaseComObject(_IOleInPlaceFrame);
                    _IOleInPlaceFrame = null;
                }
            }

            base.Dispose(disposing);
        }

        // --------------------------------------------------------------------

        public int /* IOleInPlaceFrame:: */ GetWindow([Out] out IntPtr hwnd)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::GetWindow() called");
            return _IOleInPlaceFrame.GetWindow(out hwnd);
        }

        public int /* IOleInPlaceFrame:: */ ContextSensitiveHelp([In] int fEnterMode)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::ContextSensitiveHelp() called");
            return _IOleInPlaceFrame.ContextSensitiveHelp(fEnterMode);
        }

        public int /* IOleInPlaceFrame:: */ GetBorder([Out] out IntPtr /* COMRECT */ lprectBorder)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::GetBorder() called");
            return _IOleInPlaceFrame.GetBorder(out lprectBorder);
        }

        public int /* IOleInPlaceFrame:: */ RequestBorderSpace([In] IntPtr /* COMRECT */ pborderwidths)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::RequestBorderSpace() called");
            return _IOleInPlaceFrame.RequestBorderSpace(pborderwidths);
        }

        public int /* IOleInPlaceFrame:: */ SetBorderSpace([In] IntPtr /* COMRECT */ pborderwidths)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::SetBorderSpace() called");
            return _IOleInPlaceFrame.SetBorderSpace(pborderwidths);
        }

        public int /* IOleInPlaceFrame:: */ SetActiveObject([In] IntPtr /* IOleInPlaceActiveObject */ pActiveObject, [In, MarshalAs(UnmanagedType.LPWStr)] string pszObjName)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::SetActiveObject() called");
            // need to wrap IOleInPlaceActiveObject to support this.  Used by VBE on focus. Doesn't seem to be needed by UserControl?
            //return _IOleInPlaceFrame.SetActiveObject(pActiveObject, pszObjName);
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleInPlaceFrame:: */ InsertMenus([In] IntPtr hmenuShared, [In, Out] ref IntPtr /* tagOleMenuGroupWidths */ lpMenuWidths)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::InsertMenus() called");
            return _IOleInPlaceFrame.InsertMenus(hmenuShared, lpMenuWidths);
        }

        public int /* IOleInPlaceFrame:: */ SetMenu([In] IntPtr hmenuShared, [In] IntPtr holemenu, [In] IntPtr hwndActiveObject)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::SetMenu() called");
            return _IOleInPlaceFrame.SetMenu(hmenuShared, holemenu, hwndActiveObject);
        }

        public int /* IOleInPlaceFrame:: */ RemoveMenus([In] IntPtr hmenuShared)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::RemoveMenus() called");
            return _IOleInPlaceFrame.RemoveMenus(hmenuShared);
        }

        public int /* IOleInPlaceFrame:: */ SetStatusText([In, MarshalAs(UnmanagedType.LPWStr)] string pszStatusText)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::SetStatusText() called");
            return _IOleInPlaceFrame.SetStatusText(pszStatusText);
        }

        public int /* IOleInPlaceFrame:: */ EnableModeless([In] bool fEnable)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::EnableModeless() called");
            return _IOleInPlaceFrame.EnableModeless(fEnable);
        }

        public int /* IOleInPlaceFrame:: */ TranslateAccelerator([In] IntPtr lpmsg, [In] ushort wID)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceFrame::TranslateAccelerator() called");
            return _IOleInPlaceFrame.TranslateAccelerator(lpmsg, wID);
        }
    }

    public class Wrapper_IOleClientSite : WrapperBase, COM_IOleClientSite, COM_IOleInPlaceSite, COM_IOleWindow
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        COM_IOleClientSite _IOleClientSite;        // cached object for accessing the IOleClientSite interface
        COM_IOleInPlaceSite _IOleInPlaceSite;       // cached object for accessing the IOleInPlaceSite interface
        public Wrapper_IOleInPlaceFrame _cachedFrame;           // cache the frame object returned from GetWindowContext, so that we can control the tear down

        public Wrapper_IOleClientSite(IntPtr hostObjectPtr) : base(hostObjectPtr)
        {
            if (hostObjectPtr != IntPtr.Zero)
            {
                _IOleClientSite = (COM_IOleClientSite)GetObject();
                _IOleInPlaceSite = (COM_IOleInPlaceSite)GetObject();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _cachedFrame?.Dispose();
                _cachedFrame = null;

                if (_IOleClientSite != null)
                {
                    Marshal.ReleaseComObject(_IOleClientSite);
                    _IOleClientSite = null;
                }

                if (_IOleInPlaceSite != null)
                {
                    Marshal.ReleaseComObject(_IOleInPlaceSite);
                    _IOleInPlaceSite = null;
                }
            }

            base.Dispose(disposing);
        }

        // --------------------------------------------------------------------

        public int /* IOleClientSite:: */ SaveObject()
        {
            _logger.Log(LogLevel.Trace, "IOleClientSite::SaveObject() called");
            return _IOleClientSite.SaveObject();
        }

        public int /* IOleClientSite:: */ GetMoniker([In] uint dwAssign, [In] uint dwWhichMoniker, [Out] out IntPtr /* IMoniker */ moniker)
        {
            _logger.Log(LogLevel.Trace, "IOleClientSite::GetMoniker() called");
            // need to wrap IMoniker to support this.  Not used by VBE anyway?
            //return _IOleClientSite.GetMoniker(dwAssign, dwWhichMoniker, out moniker);
            moniker = IntPtr.Zero;
            Debug.Assert(false);
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleClientSite:: */ GetContainer([Out] out IntPtr /* IOleContainer */ container)
        {
            _logger.Log(LogLevel.Trace, "IOleClientSite::GetContainer() called");
            // need to wrap IOleContainer to support this.  VBE doesn't implement this anyway (returns E_NOTIMPL)
            //return _IOleClientSite.GetContainer(out container);
            container = IntPtr.Zero;
            return (int)ComConstants.E_NOTIMPL;
        }

        public int /* IOleClientSite:: */ ShowObject()
        {
            _logger.Log(LogLevel.Trace, "IOleClientSite::ShowObject() called");
            return _IOleClientSite.ShowObject();
        }

        public int /* IOleClientSite:: */ OnShowWindow([In] int fShow)
        {
            _logger.Log(LogLevel.Trace, "IOleClientSite::OnShowWindow() called");
            return _IOleClientSite.OnShowWindow(fShow);
        }

        public int /* IOleClientSite:: */ RequestNewObjectLayout()
        {
            _logger.Log(LogLevel.Trace, "IOleClientSite::RequestNewObjectLayout() called");
            return _IOleClientSite.RequestNewObjectLayout();
        }

        // --------------------------------------------------------------------

        public int /* IOleInPlaceSite:: */ GetWindow([Out] out IntPtr hwnd)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::GetWindow() called");
            return _IOleInPlaceSite.GetWindow(out hwnd);
        }

        public int /* IOleInPlaceSite:: */ ContextSensitiveHelp([In] int fEnterMode)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::ContextSensitiveHelp() called");
            return _IOleInPlaceSite.ContextSensitiveHelp(fEnterMode);
        }

        public int /* IOleInPlaceSite:: */ CanInPlaceActivate()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::CanInPlaceActivate() called");
            return _IOleInPlaceSite.CanInPlaceActivate();
        }

        public int /* IOleInPlaceSite:: */ OnInPlaceActivate()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::OnInPlaceActivate() called");
            return _IOleInPlaceSite.OnInPlaceActivate();
        }

        public int /* IOleInPlaceSite:: */ OnUIActivate()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::OnUIActivate() called");
            return _IOleInPlaceSite.OnUIActivate();
        }

        public int /* IOleInPlaceSite:: */ GetWindowContext([Out] out IntPtr /* IOleInPlaceFrame */ ppFrame, [Out] out IntPtr /* IOleInPlaceUIWindow */ ppDoc, [Out] out IntPtr /* COMRECT */ lprcPosRect, [Out] out IntPtr /* COMRECT */ lprcClipRect, [In] IntPtr /* tagOIFI */ lpFrameInfo)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::GetWindowContext() called");
            int hr = _IOleInPlaceSite.GetWindowContext(out ppFrame, out ppDoc, out lprcPosRect, out lprcClipRect, lpFrameInfo);
            if (hr >= 0)
            {
                // call succeeded, so wrap the ppFrame with our own object so that we can control all host owned COM references
                _cachedFrame?.Dispose();
                _cachedFrame = null;
                _cachedFrame = new Wrapper_IOleInPlaceFrame(ppFrame);
                Marshal.Release(ppFrame);    // the Wrapper_IOleInPlaceFrame took its own reference, so we can release this one.
                ppFrame = _cachedFrame.CopyAggregatedReference();

                Debug.Assert(ppDoc == IntPtr.Zero);  // ppDoc not used by VBE, so no need to wrap it
            }

            return hr;
        }

        public int /* IOleInPlaceSite:: */ Scroll([In] IntPtr /* SIZE */ scrollExtant)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::Scroll() called");
            return _IOleInPlaceSite.Scroll(scrollExtant);
        }

        public int /* IOleInPlaceSite:: */ OnUIDeactivate([In] int fUndoable)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::OnUIDeactivate() called");
            return _IOleInPlaceSite.OnUIDeactivate(fUndoable);
        }

        public int /* IOleInPlaceSite:: */ OnInPlaceDeactivate()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::OnInPlaceDeactivate() called");
            return _IOleInPlaceSite.OnInPlaceDeactivate();
        }

        public int /* IOleInPlaceSite:: */ DiscardUndoState()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::DiscardUndoState() called");
            return _IOleInPlaceSite.DiscardUndoState();
        }

        public int /* IOleInPlaceSite:: */ DeactivateAndUndo()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::DeactivateAndUndo() called");
            return _IOleInPlaceSite.DeactivateAndUndo();
        }

        public int /* IOleInPlaceSite:: */ OnPosRectChange([In] IntPtr /* COMRECT */ lprcPosRect)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceSite::OnPosRectChange() called");
            return _IOleInPlaceSite.OnPosRectChange(lprcPosRect);
        }
    }
}
