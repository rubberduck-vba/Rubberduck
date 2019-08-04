using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged
{
    /// <summary>
    /// Some known COM HRESULTs used in our code
    /// </summary>
    internal enum KnownComHResults 
    {
        S_OK = 0,
        E_VBA_COMPILEERROR = unchecked((int)0x800A9C64),
        E_NOTIMPL = unchecked((int)0x80004001),
        DISP_E_EXCEPTION = unchecked((int)0x80020009),
        E_INVALIDARG = unchecked((int)0x80070057),
        TYPE_E_ELEMENTNOTFOUND = unchecked((int)0x8002802B),
    }

    /// <summary>
    /// Ensures that a wrapped COM object only responds to a specific COM interface.
    /// </summary>
    /// <typeparam name="T">The COM interface for restriction</typeparam>
    internal class RestrictComInterfaceByAggregation<T> : ICustomQueryInterface, IDisposable
    {
        private readonly IntPtr _outerObject;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="outerObject">The object that needs interface requests filtered</param>
        /// <param name="queryForType">determines whether we call QueryInterface for the interface or not</param>
        /// <remarks>if the passed in outerObject is known to point to the correct vtable for the interface, then queryForType can be false</remarks>
        /// <returns>if outerObject is IntPtr.Zero, then a null wrapper, else an aggregated wrapper</returns>
        public RestrictComInterfaceByAggregation(IntPtr outerObject, bool queryForType = true)
        {
            if (queryForType)
            {
                var iid = typeof(T).GUID;
                if (ComHelper.HRESULT_FAILED(RdMarshal.QueryInterface(outerObject, ref iid, out _outerObject)))
                {
                    // allow null wrapping here
                    return;
                }
            }
            else
            {
                _outerObject = outerObject;
                RdMarshal.AddRef(_outerObject);
            }

            var clrAggregator = RdMarshal.CreateAggregatedObject(_outerObject, this);
            WrappedObject = (T)RdMarshal.GetObjectForIUnknown(clrAggregator);        // when this CCW object gets released, it will free the aggObjInner (well, after GC)
            RdMarshal.Release(clrAggregator);         // _wrappedObject holds a reference to this now
        }

        public T WrappedObject { get; set; }

        // this extracts the wrapped object.  caller is then responsible for calling Marshal.ReleaseComObject on it.
        public T ExtractWrappedObject()
        {
            var retVal = WrappedObject;
            WrappedObject = default;
            return retVal;
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
            _isDisposed = true;

            if (WrappedObject != null)
            {
                RdMarshal.ReleaseComObject(WrappedObject);
            }

            if (_outerObject != IntPtr.Zero)
            {
                RdMarshal.Release(_outerObject);

                // dont set _outerObject to IntPtr.Zero here, as GetInterface() can still be called by the outer RCW
                // if it is still alive. For example, if ExtractWrappedObject was used and the outer object hasn't yet
                // been released with ReleaseComObject.  In that circumstance _outerObject will still be a valid pointer  
                // due to the internally held reference, and so GetInterface() calls past this point are still OK.
            }
        }

        // The aggregation magic starts here
        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;

            if (iid == typeof(T).GUID)
            {
                ppv = _outerObject;
                RdMarshal.AddRef(_outerObject);
                return CustomQueryInterfaceResult.Handled;
            }

            return CustomQueryInterfaceResult.Failed;
        }
    }

    /// <summary>
    /// Exposes some special routines for dealing with COM interop
    /// </summary>
    internal static class ComHelper
    {
        /// <summary>
        /// Equivalent of the Windows FAILED() macro in C
        /// see https://msdn.microsoft.com/en-us/library/windows/desktop/ms693474(v=vs.85).aspx
        /// </summary>
        /// <param name="hr">HRESULT from a COM API call</param>
        /// <returns>true if the HRESULT indicated failure</returns>
        public static bool HRESULT_FAILED(int hr) => hr < 0;

        // simply check if a COM object supports a particular COM interface
        // (without doing any casting by the CLR, which does much more than this under the covers)
        public static bool DoesComObjPtrSupportInterface<T>(IntPtr comObjPtr)
        {
            var iid = typeof(T).GUID;
            var hr = RdMarshal.QueryInterface(comObjPtr, ref iid, out var outInterfacePtr);
            if (!ComHelper.HRESULT_FAILED(hr))
            {
                RdMarshal.Release(outInterfacePtr);
                return true;
            }
            return false;
        }

        // ComCastViaAggregation creates an aggregated object to forcibly cast a pointer
        // to a specific interface.  The aggregator is disposed before returning.  Caller must release the
        // returned RCW with Marshal.ReleaseComObject.  
        public static T ComCastViaAggregation<T>(IntPtr rawObjectPtr, bool queryForType = true)
        {
            using (var aggregator = new RestrictComInterfaceByAggregation<T>(rawObjectPtr, queryForType))
            {
                return aggregator.ExtractWrappedObject();
            }
        }
    }
}
