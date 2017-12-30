using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.WindowsApi;
using User32 = Rubberduck.Common.WinAPI.User32;
using NLog;

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

    public enum COMConstants
    {
        E_NOTIMPL = -2147467263,
        S_OK = 0
    }   
  
    [ComImport(), Guid("00000112-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleObject
    {
        [PreserveSig] int SetClientSite         ([In] IntPtr /* IOleClientSite */ pClientSite);
        [PreserveSig] int GetClientSite         ([Out] out IntPtr /* IOleClientSite */ ppClientSite);
        [PreserveSig] int SetHostNames          ([In, MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [In, MarshalAs(UnmanagedType.LPWStr)] string szContainerObj);
        [PreserveSig] int Close                 ([In] uint dwSaveOption);
        [PreserveSig] int SetMoniker            ([In] uint dwWhichMoniker, [In] IntPtr /* IMoniker */ pmk);
        [PreserveSig] int GetMoniker            ([In] uint dwAssign, [In] uint dwWhichMoniker, [Out] out IntPtr /* IMoniker */ ppmk);
        [PreserveSig] int InitFromData          ([In] IntPtr /* IDataObject */ pDataObject, [In] int fCreation, [In] uint dwReserved);
        [PreserveSig] int GetClipboardData      ([In] uint dwReserved, [Out] out IntPtr /*IDataObject*/ ppDataObject);
        [PreserveSig] int DoVerb                ([In] int iVerb, [In] IntPtr /* MSG, nullable ref */ lpmsg, [In] IntPtr /* IOleClientSite */ pActiveSite, [In] int lindex, [In] IntPtr hwndParent, [In] IntPtr /* COMRECT */ lprcPosRect);
        [PreserveSig] int EnumVerbs             ([Out] out IntPtr /* IEnumOLEVERB */ ppEnumOleVerb);
        [PreserveSig] int Update                ();
        [PreserveSig] int IsUpToDate            ();
        [PreserveSig] int GetUserClassID        ([Out] out Guid pClsid);
        [PreserveSig] int GetUserType           ([In] uint dwFormOfType, [Out, MarshalAs(UnmanagedType.LPWStr)] out string pszUserType);
        [PreserveSig] int SetExtent             ([In] uint dwDrawAspect, [In] IntPtr /* tagSIZE */ psizel);
        [PreserveSig] int GetExtent             ([In] uint dwDrawAspect, [Out] out IntPtr /* tagSIZE */ psizel);
        [PreserveSig] int Advise                ([In] IntPtr /* IAdviseSink */ pAdvSink, [Out] out uint pdwConnection);
        [PreserveSig] int Unadvise              ([In] uint pdwConnection);
        [PreserveSig] int EnumAdvise            ([Out] out IntPtr /* IEnumSTATDATA */ enumAdvise);
        [PreserveSig] int GetMiscStatus         ([In] uint dwAspect, [Out] out uint pdwStatus);
        [PreserveSig] int SetColorScheme        ([In] IntPtr /* tagLOGPALETTE */ pLogpal);
    };

    [ComImport(), Guid("00000113-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleInPlaceObject /* : COM_IOleWindow */
    {
        [PreserveSig] int GetWindow([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp([In] int fEnterMode);
        [PreserveSig] int InPlaceDeactivate();
        [PreserveSig] int UIDeactivate();
        [PreserveSig] int SetObjectRects([In] IntPtr /* COMRECT */ lprcPosRect, [In] IntPtr /* COMRECT */ lprcClipRect);
        [PreserveSig] int ReactivateAndUndo();
    }

    [ComImport(), Guid("00000118-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleClientSite
    { 
        [PreserveSig] int SaveObject            ();
        [PreserveSig] int GetMoniker            ([In] uint dwAssign, [In] uint dwWhichMoniker, [Out] out IntPtr /* IMoniker */ moniker);
        [PreserveSig] int GetContainer          ([Out] out IntPtr /* IOleContainer */ container);
        [PreserveSig] int ShowObject            ();
        [PreserveSig] int OnShowWindow          ([In] int fShow);
        [PreserveSig] int RequestNewObjectLayout();
    }

    [ComImport(), Guid("00000114-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleWindow
    {
        [PreserveSig] int GetWindow             ([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp  ([In] int fEnterMode);
    }

    [ComImport(), Guid("00000115-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleInPlaceUIWindow /* : COM_IOleWindow */
    {
        [PreserveSig] int GetWindow             ([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp  ([In] int fEnterMode);
        [PreserveSig] int GetBorder             ([Out] out IntPtr /* COMRECT */ lprectBorder);
        [PreserveSig] int RequestBorderSpace    ([In] IntPtr /* COMRECT */ pborderwidths);
        [PreserveSig] int SetBorderSpace        ([In] IntPtr /* COMRECT */ pborderwidths);
        [PreserveSig] int SetActiveObject       ([In] IntPtr /* IOleInPlaceActiveObject */ pActiveObject, [In, MarshalAs(UnmanagedType.LPWStr)] string pszObjName);
    }

    [ComImport(), Guid("00000116-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleInPlaceFrame /* : COM_IOleInPlaceUIWindow */
    {
        [PreserveSig] int GetWindow             ([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp  ([In] int fEnterMode);
        [PreserveSig] int GetBorder             ([Out] out IntPtr /* COMRECT */ lprectBorder);
        [PreserveSig] int RequestBorderSpace    ([In] IntPtr /* COMRECT */ pborderwidths);
        [PreserveSig] int SetBorderSpace        ([In] IntPtr /* COMRECT */ pborderwidths);
        [PreserveSig] int SetActiveObject       ([In] IntPtr /* IOleInPlaceActiveObject */ pActiveObject, [In, MarshalAs(UnmanagedType.LPWStr)] string pszObjName);
        [PreserveSig] int InsertMenus           ([In] IntPtr hmenuShared, [In, Out] ref IntPtr /* tagOleMenuGroupWidths */ lpMenuWidths);
        [PreserveSig] int SetMenu               ([In] IntPtr hmenuShared, [In] IntPtr holemenu, [In] IntPtr hwndActiveObject);
        [PreserveSig] int RemoveMenus           ([In] IntPtr hmenuShared);
        [PreserveSig] int SetStatusText         ([In, MarshalAs(UnmanagedType.LPWStr)] string pszStatusText);
        [PreserveSig] int EnableModeless        ([In] bool fEnable);
        [PreserveSig] int TranslateAccelerator  ([In] IntPtr lpmsg, [In] ushort wID);
    }

    [ComImport(), Guid("00000119-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface COM_IOleInPlaceSite /* : COM_IOleWindow */
    {
        [PreserveSig] int GetWindow([Out] out IntPtr hwnd);
        [PreserveSig] int ContextSensitiveHelp([In] int fEnterMode);
        [PreserveSig] int CanInPlaceActivate    ();
        [PreserveSig] int OnInPlaceActivate     ();
        [PreserveSig] int OnUIActivate          ();
        [PreserveSig] int GetWindowContext      ([Out] out IntPtr /* IOleInPlaceFrame */ ppFrame, [Out] out IntPtr /* IOleInPlaceUIWindow */ ppDoc, [Out] out IntPtr /* COMRECT */ lprcPosRect, [Out] out IntPtr /* COMRECT */ lprcClipRect, [In] IntPtr /* tagOIFI */ lpFrameInfo);
        [PreserveSig] int Scroll                ([In] IntPtr /* tagSIZE */ scrollExtant);
        [PreserveSig] int OnUIDeactivate        ([In] int fUndoable);
        [PreserveSig] int OnInPlaceDeactivate   ();
        [PreserveSig] int DiscardUndoState      ();
        [PreserveSig] int DeactivateAndUndo     ();
        [PreserveSig] int OnPosRectChange       ([In] IntPtr /* COMRECT */ lprcPosRect);
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
            _outerObject = null;
            _supportedTypes = null;
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

        public virtual void Dispose()
        {
            if (_aggregatedObjectPtr != IntPtr.Zero)
            {
                Marshal.Release(_aggregatedObjectPtr);
                _aggregatedObjectPtr = IntPtr.Zero;
            }
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

        public override void Dispose()
        {
            if (_hostObjectPtr != IntPtr.Zero)
            {
                Marshal.Release(_hostObjectPtr);
                _hostObjectPtr = IntPtr.Zero;
            }

            base.Dispose();
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

        public override void Dispose()
        {
            if (_IOleInPlaceFrame != null)
            {
                Marshal.ReleaseComObject(_IOleInPlaceFrame);
                _IOleInPlaceFrame = null;
            }
            base.Dispose();
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
            return (int)COMConstants.E_NOTIMPL;
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
        private readonly Logger         _logger = LogManager.GetCurrentClassLogger();

        COM_IOleClientSite              _IOleClientSite;        // cached object for accessing the IOleClientSite interface
        COM_IOleInPlaceSite             _IOleInPlaceSite;       // cached object for accessing the IOleInPlaceSite interface
        public Wrapper_IOleInPlaceFrame _cachedFrame;           // cache the frame object returned from GetWindowContext, so that we can control the tear down

        public Wrapper_IOleClientSite(IntPtr hostObjectPtr) : base(hostObjectPtr)
        {
            if (hostObjectPtr != IntPtr.Zero)
            {
                _IOleClientSite = (COM_IOleClientSite)GetObject();
                _IOleInPlaceSite = (COM_IOleInPlaceSite)GetObject();
            }
        }

        public override void Dispose()
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
           
            base.Dispose();
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
            return (int)COMConstants.E_NOTIMPL;
        }

        public int /* IOleClientSite:: */ GetContainer([Out] out IntPtr /* IOleContainer */ container)
        {
            _logger.Log(LogLevel.Trace, "IOleClientSite::GetContainer() called");
            // need to wrap IOleContainer to support this.  VBE doesn't implement this anyway (returns E_NOTIMPL)
            //return _IOleClientSite.GetContainer(out container);
            container = IntPtr.Zero;
            return (int)COMConstants.E_NOTIMPL;
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

    // ExposedUserControl - wrapper for UserControl that also exposes the underlying 
    // IOleObject and IOleInPlaceObject COM interfaces implemented by it
    public class ExposedUserControl : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public COM_IOleObject _IOleObject;                  // cached interface obtained from UserConrol
        public COM_IOleInPlaceObject _IOleInPlaceObject;    // cached interface obtained from UserConrol

        public ExposedUserControl()
        {
            _logger.Log(LogLevel.Trace, "ExposedUserControl constructor called");
            
            // Gain access to the IOleObject and IOleInPlaceObject interfaces implemented by the UserControl
            _IOleObject = (COM_IOleObject)AggregationHelper.ObtainInternalInterface(this, GetType().GetInterface("IOleObject"));
            _IOleInPlaceObject = (COM_IOleInPlaceObject)AggregationHelper.ObtainInternalInterface(this, GetType().GetInterface("IOleInPlaceObject"));
        }

        protected override void Dispose(bool disposing)
        {
            if (_IOleObject != null)
            {
                Marshal.ReleaseComObject(_IOleObject);
                _IOleObject = null;
            }

            if (_IOleInPlaceObject != null)
            {
                Marshal.ReleaseComObject(_IOleInPlaceObject);
                _IOleInPlaceObject = null;
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

    [ComVisible(true)]
    [Guid(RubberduckGuid.DockableWindowHostGuid)]
    [ProgId(RubberduckProgId.DockableWindowHostProgId)]
    [EditorBrowsable(EditorBrowsableState.Never)]
    //Nothing breaks because we declare a ProgId
    // ReSharper disable once InconsistentNaming
    //Underscores make classes invisible to VB6 object explorer
    public class _DockableWindowHost : COM_IOleObject, COM_IOleInPlaceObject, COM_IOleWindow
    {
        public static string RegisteredProgId => RubberduckProgId.DockableWindowHostProgId;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private ExposedUserControl _userControl = new ExposedUserControl();
        private Wrapper_IOleClientSite _cachedClientSite;

        public void Release()
        {
            // WARNING: Disposal of _userControl / _cachedClientSite should be handled in IOleObject::Close(), not here, see top comments
            RemoveChildControlsFromExposedControl();
        }

        private void RemoveChildControlsFromExposedControl()
        {
            foreach(UserControl control in _userControl.Controls)
            {
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
                return _userControl._IOleObject.SetClientSite(_cachedClientSite.PeekAggregatedReference());     // callee will take its own reference
            }
            return (int)COMConstants.S_OK;
        }

        public int /* IOleObject:: */ GetClientSite([Out] out IntPtr /* IOleClientSite */ ppClientSite)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetClientSite() called");
            ppClientSite = _cachedClientSite?.CopyAggregatedReference() ?? IntPtr.Zero;
            return (int)COMConstants.S_OK;
        }

        public int /* IOleObject:: */ SetHostNames([In, MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [In, MarshalAs(UnmanagedType.LPWStr)] string szContainerObj)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::SetHostNames() called");
            return _userControl._IOleObject.SetHostNames(szContainerApp, szContainerObj);
        }

        public int /* IOleObject:: */ Close([In] uint dwSaveOption)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Close() called");
            int hr = _userControl._IOleObject.Close(dwSaveOption);

            // IOleObject::SetClientSite is typically called with pClientSite = null just before calling IOleObject::Close()
            // If it didn't, we release all host COM objects here instead,

            // This is the point where we can deterministically, and safely release our COM references for this ActiveX control.
            // Moreover, we can release the UserControl COM references, as Close() should be the very last call into the IOleObject interface.
            PerformUserControlShutdown();

            return hr;
        }

        private void PerformUserControlShutdown()
        {
            ReleaseCOMReferenceOfSctiveXControl();
            ReleasedExposedControl();
            UnsubclassParent();

            GC.Collect(); //todo: Release enough COM objects to make this not necessary anymore.
            GC.WaitForPendingFinalizers();
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

        private void ReleaseCOMReferenceOfSctiveXControl()
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
            return (int)COMConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ GetMoniker([In] uint dwAssign, [In] uint dwWhichMoniker, [Out] out IntPtr /* IMoniker */ ppmk)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetMoniker() called");
            // need to wrap IMoniker to support this.  Not used by VBE anyway?
            //return _IOleObject.GetMoniker(dwAssign, dwWhichMoniker, out ppmk);
            ppmk = IntPtr.Zero;
            Debug.Assert(false);
            return (int)COMConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ InitFromData([In] IntPtr /* IDataObject */ pDataObject, [In] int fCreation, [In] uint dwReserved)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::InitFromData() called");
            // need to wrap IDataObject to support this.  Not used by VBE anyway?
            //return _IOleObject.InitFromData(pDataObject, fCreation, dwReserved);
            Debug.Assert(false);
            return (int)COMConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ GetClipboardData([In] uint dwReserved, [Out] out IntPtr /*IDataObject*/ ppDataObject)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetClipboardData() called");
            // need to wrap IDataObject to support this.  Not used by VBE anyway?
            //return _IOleObject.GetClipboardData(dwReserved, out ppDataObject);
            ppDataObject = IntPtr.Zero;
            Debug.Assert(false);
            return (int)COMConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ DoVerb([In] int iVerb, [In] IntPtr lpmsg, [In] IntPtr /* IOleClientSite */ pActiveSite, [In] int lindex, [In] IntPtr hwndParent, [In] IntPtr /* COMRECT */ lprcPosRect)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::DoVerb() called");
            // pActiveSite is not used by the UserControl implementation.  Either wrap it or pass null instead
            pActiveSite = IntPtr.Zero;
            return _userControl._IOleObject.DoVerb(iVerb, lpmsg, pActiveSite, lindex, hwndParent, lprcPosRect);
        }

        public int /* IOleObject:: */ EnumVerbs([Out] out IntPtr /* IEnumOLEVERB */ ppEnumOleVerb)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::EnumVerbs() called");
            // need to wrap IEnumOLEVERB to support this.  Not used by VBE anyway?
            //return _IOleObject.EnumVerbs(out ppEnumOleVerb);
            ppEnumOleVerb = IntPtr.Zero;
            Debug.Assert(false);
            return (int)COMConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ Update()
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Update() called");
            return _userControl._IOleObject.Update();
        }

        public int /* IOleObject:: */ IsUpToDate()
        {
            _logger.Log(LogLevel.Trace, "IOleObject::IsUpToDate() called");
            return _userControl._IOleObject.IsUpToDate();
        }

        public int /* IOleObject:: */ GetUserClassID([Out] out Guid pClsid)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetUserClassID() called");
            return _userControl._IOleObject.GetUserClassID(out pClsid);
        }

        public int /* IOleObject:: */ GetUserType([In] uint dwFormOfType, [Out, MarshalAs(UnmanagedType.LPWStr)] out string pszUserType)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetUserType() called");
            return _userControl._IOleObject.GetUserType(dwFormOfType, out pszUserType);
        }

        public int /* IOleObject:: */ SetExtent([In] uint dwDrawAspect, [In] IntPtr /* tagSIZE */ psizel)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::SetExtent() called");
            return _userControl._IOleObject.SetExtent(dwDrawAspect, psizel);
        }

        public int /* IOleObject:: */ GetExtent([In] uint dwDrawAspect, [Out] out IntPtr /* tagSIZE */ psizel)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetExtent() called");
            return _userControl._IOleObject.GetExtent(dwDrawAspect, out psizel);
        }

        public int /* IOleObject:: */ Advise([In] IntPtr /* IAdviseSink */ pAdvSink, [Out] out uint pdwConnection)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Advise() called");
            // need to wrap IAdviseSink to support this. VBE does try to use this, but the events don't look interesting?
            pdwConnection = 0;
            return (int)COMConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ Unadvise([In] uint pdwConnection)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::Unadvise() called");
            // No sense supporting Unadvise, as we're not supporting Advise
            //return _IOleObject.Unadvise(pdwConnection);
            //Debug.Assert(false);                              stupid VBE still calls us, despite us not implementing Advise()
            return (int)COMConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ EnumAdvise([Out] out IntPtr /* IEnumSTATDATA */ enumAdvise)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::EnumAdvise() called");
            // need to wrap IEnumSTATDATA to support this. No sense supporting EnumAdvise, as we're not supporting Advise
            //return _IOleObject.EnumAdvise(out enumAdvise);
            enumAdvise = IntPtr.Zero;
            return (int)COMConstants.E_NOTIMPL;
        }

        public int /* IOleObject:: */ GetMiscStatus([In] uint dwAspect, [Out] out uint pdwStatus)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::GetMiscStatus() called");
            return _userControl._IOleObject.GetMiscStatus(dwAspect, out pdwStatus);
        }

        public int /* IOleObject:: */ SetColorScheme([In] IntPtr /* tagLOGPALETTE */ pLogpal)
        {
            _logger.Log(LogLevel.Trace, "IOleObject::SetColorScheme() called");
            return _userControl._IOleObject.SetColorScheme(pLogpal);
        }

        // --------------------------------------------------------------------

        public int /* IOleInPlaceObject:: */ GetWindow([Out] out IntPtr hwnd)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::GetWindow() called");
            return _userControl._IOleInPlaceObject.GetWindow(out hwnd);
        }

        public int /* IOleInPlaceObject:: */ ContextSensitiveHelp([In] int fEnterMode)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::ContextSensitiveHelp() called");
            return _userControl._IOleInPlaceObject.ContextSensitiveHelp(fEnterMode);
        }

        public int /* IOleInPlaceObject:: */ InPlaceDeactivate()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::InPlaceDeactivate() called");
            return _userControl._IOleInPlaceObject.InPlaceDeactivate();
        }

        public int /* IOleInPlaceObject:: */ UIDeactivate()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::UIDeactivate() called");
            return _userControl._IOleInPlaceObject.UIDeactivate();
        }

        public int /* IOleInPlaceObject:: */ SetObjectRects([In] IntPtr lprcPosRect, [In] IntPtr lprcClipRect)
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::SetObjectRects() called");
            return _userControl._IOleInPlaceObject.SetObjectRects(lprcPosRect, lprcClipRect);
        }

        public int /* IOleInPlaceObject:: */ ReactivateAndUndo()
        {
            _logger.Log(LogLevel.Trace, "IOleInPlaceObject::ReactivateAndUndo() called");
            return _userControl._IOleInPlaceObject.ReactivateAndUndo();
        }

        // old stuff from old _DockableWindowHost --------------------------------------- [START]

        private void OnCallBackEvent(object sender, SubClassingWindowEventArgs e)
        {
            if (e.Closing)
            {
                return;
            }
            var param = new LParam { Value = (uint)e.LParam };
            _userControl.Size = new Size(param.LowWord, param.HighWord);
        }

        internal void AddUserControl(UserControl control, IntPtr vbeHwnd)
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

        [ComVisible(false)]
        public class ParentWindow : SubclassingWindow
        {
            public event SubClassingWindowEventHandler CallBackEvent;
            public delegate void SubClassingWindowEventHandler(object sender, SubClassingWindowEventArgs e);

            private readonly IntPtr _vbeHwnd;

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
    }
}
