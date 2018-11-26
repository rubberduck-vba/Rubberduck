using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    // The CLR automagically handles COM enumeration by custom marshalling IEnumVARIANT to a managed class (see EnumeratorToEnumVariantMarshaler)
    // But as we need explicit control over all COM objects in RD, this is unacceptable.  We need to obtain the explicit IEnumVARIANT interface
    // and ensure this RCW is destroyed in a timely fashion, using Marshal.ReleaseComObject.
    // The automatic custom marshalling of the enumeration getter method (DISPID_ENUM) prohibits access to the underlying IEnumVARIANT interface.
    // To work around it, we must call the IDispatch:::Invoke method directly (instead of using CLRs normal late-bound method calling ability).

    [ComImport(), Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]        
    internal interface IDispatch
    {
        [PreserveSig] int GetTypeInfoCount([Out] out uint pctinfo);
        [PreserveSig] int GetTypeInfo([In] uint iTInfo, [In] uint lcid, [Out] out ComTypes.ITypeInfo pTypeInfo);
        [PreserveSig] int GetIDsOfNames([In] ref Guid riid, [In] string[] rgszNames, [In] uint cNames, [In] uint lcid, [Out] out int[] rgDispId);
 
        [PreserveSig]
        int Invoke([In] int dispIdMember,
            [In] ref Guid riid,
            [In] uint lcid,
            [In] uint dwFlags,
            [In, Out] ref ComTypes.DISPPARAMS pDispParams,
            [Out] out Object pVarResult,
            [In, Out] ref ComTypes.EXCEPINFO pExcepInfo,
            [Out] out uint pArgErr);
    }

    internal class IDispatchHelper
    {
        public enum StandardDispIds : int
        {
            DISPID_ENUM = -4
        }

        public enum InvokeKind : int
        {
            DISPATCH_METHOD = 1,
            DISPATCH_PROPERTYGET = 2,
            DISPATCH_PROPERTYPUT = 4,
            DISPATCH_PROPERTYPUTREF = 8,
        }
        
        public static object PropertyGet_NoArgs(IDispatch obj, int memberId)
        {
            var pDispParams = new ComTypes.DISPPARAMS();
            var pExcepInfo = new ComTypes.EXCEPINFO();
            Guid guid = new Guid();

            int hr = obj.Invoke(memberId, ref guid, 0, (uint)(InvokeKind.DISPATCH_METHOD | InvokeKind.DISPATCH_PROPERTYGET), 
                                    ref pDispParams, out var pVarResult, ref pExcepInfo, out uint ErrArg);

            if (hr < 0)
            {
                // could expand this to better handle DISP_E_EXCEPTION
                throw Marshal.GetExceptionForHR(hr);
            }

            return pVarResult;
        }
    }

    [ComImport(), Guid("00020404-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumVARIANT
    {
        // rgVar is technically an unmanaged array here, but we only ever call with celt=1, so this is compatible.
        [PreserveSig] int Next([In] uint celt, [Out] out object rgVar, [Out] out uint pceltFetched);

        [PreserveSig] int Skip([In] uint celt);
        [PreserveSig] int Reset();
        [PreserveSig] int Clone([Out] out IEnumVARIANT retval);
    }

    public sealed class ComWrapperEnumerator<TWrapperItem> : IEnumerator<TWrapperItem>
        where TWrapperItem : class
    {
        private readonly Func<object, TWrapperItem> _itemWrapper;
        private readonly IEnumVARIANT _enumeratorRCW;
        private TWrapperItem _currentItem;

        public ComWrapperEnumerator(object source, Func<object, TWrapperItem> itemWrapper)
        {
            _itemWrapper = itemWrapper;

            if (source != null)
            {
                _enumeratorRCW = (IEnumVARIANT)IDispatchHelper.PropertyGet_NoArgs((IDispatch)source, (int)IDispatchHelper.StandardDispIds.DISPID_ENUM);
                ((IEnumerator)this).Reset();  // precaution 
            }
        }

        void IEnumerator.Reset()
        {
            if (!IsWrappingNullReference)
            {
                int hr = _enumeratorRCW.Reset();      
                if (hr < 0)
                {
                    throw Marshal.GetExceptionForHR(hr);
                }
            }
        }
        
        public bool IsWrappingNullReference => _enumeratorRCW == null;
        
        public TWrapperItem Current => _currentItem; 
        object IEnumerator.Current => _currentItem;

        bool IEnumerator.MoveNext()
        {
            if (IsWrappingNullReference)
            {
                return false;
            }

            _currentItem = null;

            object currentItemRCW;
            uint celtFetched;
            int hr = _enumeratorRCW.Next(1, out currentItemRCW, out celtFetched);
            // hr == S_FALSE (1) or S_OK (0), or <0 means error

            _currentItem = _itemWrapper.Invoke(currentItemRCW);     // creates a null wrapped reference even on end/error, just as a precaution
            
            if (hr < 0)
            {
                throw Marshal.GetExceptionForHR(hr);
            }
           
            return (celtFetched == 1);      // celtFetched will be 0 when we reach the end of the collection
        }

        public void Dispose()
        {
            if (!IsWrappingNullReference)
            {
                Marshal.ReleaseComObject(_enumeratorRCW);
            }
        }
    }
}