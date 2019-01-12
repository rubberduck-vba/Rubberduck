using System;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// A compatible version of ITypeInfo, where COM objects are outputted as IntPtrs instead of objects
    /// see https://msdn.microsoft.com/en-gb/library/windows/desktop/ms221696(v=vs.85).aspx
    /// </summary>
    [ComImport(), Guid("00020401-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITypeInfoInternal
    {
        [PreserveSig] /*HRESULT*/ int GetTypeAttr(/*out TYPEATTR*/ IntPtr ppTypeAttr);
        [PreserveSig] /*HRESULT*/ int GetTypeComp(/*out ITypeComp*/ IntPtr ppTComp);
        [PreserveSig] /*HRESULT*/ int GetFuncDesc(int index, /*out FUNCDESC*/ IntPtr ppFuncDesc);
        [PreserveSig] /*HRESULT*/ int GetVarDesc(int index, /*out VARDESC*/ IntPtr ppVarDesc);
        [PreserveSig] /*HRESULT*/ int GetNames(int memid, /*string[]*/ IntPtr rgBstrNames, int cMaxNames, /*out int*/ IntPtr pcNames);
        [PreserveSig] /*HRESULT*/ int GetRefTypeOfImplType(int index, /*out HREFTYPE*/ IntPtr href);
        [PreserveSig] /*HRESULT*/ int GetImplTypeFlags(int index, /*out ComTypes.IMPLTYPEFLAGS*/ IntPtr pImplTypeFlags);
        [PreserveSig] /*HRESULT*/ int GetIDsOfNames(/*string[]*/ IntPtr rgszNames, int cNames, /*out MEMBERID*/ IntPtr pMemId);
        [PreserveSig] /*HRESULT*/ int Invoke(/*object*/ IntPtr pvInstance, int memid, short wFlags,/*ref ComTypes.DISPPARAMS*/ IntPtr pDispParams, /*out VARIANT*/ IntPtr pVarResult, /*out EXCEPINFO*/ IntPtr pExcepInfo, /*out int*/ IntPtr puArgErr);
        [PreserveSig] /*HRESULT*/ int GetDocumentation(int index, /*out string*/ IntPtr strName, /*out string*/ IntPtr strDocString, /*out int*/ IntPtr dwHelpContext, /*out string*/ IntPtr strHelpFile);
        [PreserveSig] /*HRESULT*/ int GetDllEntry(int memid, ComTypes.INVOKEKIND invKind, /*out string*/ IntPtr pBstrDllName, /*out string*/IntPtr pBstrName, /*out short*/ IntPtr pwOrdinal);
        [PreserveSig] /*HRESULT*/ int GetRefTypeInfo(int hRef, /*out ITypeInfo*/ IntPtr ppTI);
        [PreserveSig] /*HRESULT*/ int AddressOfMember(int memid, ComTypes.INVOKEKIND invKind, /*out IntPtr*/ IntPtr ppv);
        [PreserveSig] /*HRESULT*/ int CreateInstance(/*object*/ IntPtr pUnkOuter, ref Guid riid, /*out IntPtr*/ IntPtr ppvObj);
        [PreserveSig] /*HRESULT*/ int GetMops(int memid, /*out string*/ IntPtr pBstrMops);
        [PreserveSig] /*HRESULT*/ int GetContainingTypeLib(/*out ITypeLib*/ IntPtr ppTLB, /*out int*/ IntPtr pIndex);
        [PreserveSig] void ReleaseTypeAttr(/*TYPEATTR*/ IntPtr pTypeAttr);
        [PreserveSig] void ReleaseFuncDesc(/*FUNCDESC*/ IntPtr pFuncDesc);
        [PreserveSig] void ReleaseVarDesc(/*VARDESC*/ IntPtr pVarDesc);
    }

    // This class marshals ComTypes.ITypeLib members to ITypeLibInternal
    // ITypeLibInternal must be inherited BEFORE ComTypes.ITypeLib as they both have the same IID 
    // this will ensure QueryInterface(IID_ITypeLib) returns ITypeLibInternal, not ComTypes.ITypeLib
    public abstract class ITypeInfoInternalSelfMarshalForwarder : ITypeInfoInternal, ComTypes.ITypeInfo
    {
        private ITypeInfoInternal _this_Internal => (ITypeInfoInternal)this;

        private void HandleBadHRESULT(int hr)
        {
            throw Marshal.GetExceptionForHR(hr);
        }

        void ComTypes.ITypeInfo.GetContainingTypeLib(out ComTypes.ITypeLib ppTLB, out int pIndex)
        {
            // initialize out parameters
            ppTLB = null;
            pIndex = 0;

            using (var typeLibPtr = AddressableVariables.CreateObjectPtr<ComTypes.ITypeLib>())
            using (var indexPtr = AddressableVariables.Create<int>())
            {
                int hr = _this_Internal.GetContainingTypeLib(typeLibPtr.Address, indexPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                ppTLB = typeLibPtr.Value;
                pIndex = indexPtr.Value;
            }
        }

        void ComTypes.ITypeInfo.GetTypeAttr(out IntPtr ppTypeAttr)
        {
            // initialize out parameters
            ppTypeAttr = IntPtr.Zero;

            using (var typeAttrPtr = AddressableVariables.CreatePtrTo<ComTypes.TYPEATTR>())
            {
                var hr = _this_Internal.GetTypeAttr(typeAttrPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                ppTypeAttr = typeAttrPtr.Value.Address;        // dereference the ptr, and take the contents address
            }
        }

        void ComTypes.ITypeInfo.GetTypeComp(out ComTypes.ITypeComp ppTComp)
        {
            // initialize out parameters
            ppTComp = null;

            using (var typeCompPtr = AddressableVariables.CreateObjectPtr<ComTypes.ITypeComp>())
            {
                var hr = _this_Internal.GetTypeComp(typeCompPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                ppTComp = typeCompPtr.Value;
            }
        }

        void ComTypes.ITypeInfo.GetFuncDesc(int index, out IntPtr ppFuncDesc)
        {
            // initialize out parameters
            ppFuncDesc = IntPtr.Zero;

            using (var funcDescPtr = AddressableVariables.CreatePtrTo<ComTypes.FUNCDESC>())
            {
                var hr = _this_Internal.GetFuncDesc(index, funcDescPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                ppFuncDesc = funcDescPtr.Value.Address;        // dereference the ptr, and take the contents address
            }
        }

        void ComTypes.ITypeInfo.GetVarDesc(int index, out IntPtr ppVarDesc)
        {
            // initialize out parameters
            ppVarDesc = IntPtr.Zero;

            using (var varDescPtr = AddressableVariables.CreatePtrTo<ComTypes.VARDESC>())
            {
                var hr = _this_Internal.GetVarDesc(index, varDescPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                ppVarDesc = varDescPtr.Value.Address;          // dereference the ptr, and take the contents address
            }
        }

        void ComTypes.ITypeInfo.GetNames(int memid, string[] rgBstrNames, int cMaxNames, out int pcNames)
        {
            // initialize out parameters
            pcNames = 0;

            using (var names = AddressableVariables.CreateBSTR(cMaxNames))
            using (var namesCount = AddressableVariables.Create<int>())
            {
                var hr = _this_Internal.GetNames(memid, names.Address, cMaxNames, namesCount.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                names.CopyArrayTo(rgBstrNames);
                pcNames = namesCount.Value;
            }
        }

        void ComTypes.ITypeInfo.GetRefTypeOfImplType(int index, out int href)
        {
            // initialize out parameters
            href = 0;

            using (var outHref = AddressableVariables.Create<int>())
            {
                int hr = _this_Internal.GetRefTypeOfImplType(index, outHref.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                href = outHref.Value;
            }
        }

        void ComTypes.ITypeInfo.GetImplTypeFlags(int index, out ComTypes.IMPLTYPEFLAGS pImplTypeFlags)
        {
            // initialize out parameters
            pImplTypeFlags = 0;

            using (var implTypeFlags = AddressableVariables.Create<ComTypes.IMPLTYPEFLAGS>())
            {
                int hr = _this_Internal.GetImplTypeFlags(index, implTypeFlags.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                pImplTypeFlags = implTypeFlags.Value;
            }
        }

        void ComTypes.ITypeInfo.GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId)
        {
            // We can't use the managed arrays as passed in.  We create our own unmanaged arrays, 
            // and copy them into the managed ones on completion
            using (var names = AddressableVariables.CreateBSTR(cNames))
            using (var memberIds = AddressableVariables.Create<int>(cNames))
            {
                var hr = _this_Internal.GetIDsOfNames(names.Address, cNames, memberIds.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                names.CopyArrayTo(rgszNames);
                memberIds.CopyArrayTo(pMemId);
            }
        }

        void ComTypes.ITypeInfo.Invoke(object pvInstance, int memid, short wFlags, ref ComTypes.DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr)
        {
            // with this having in-out parameters, it would be difficult to forward it to the unfriendly implementation
            // let's not bother, since this isn't even implemented by VBA typeinfos anyway.
            throw new NotImplementedException();
        }

        void ComTypes.ITypeInfo.GetDocumentation(int memid, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile)
        {
            // initialize out parameters
            strName = null;
            strDocString = null;
            dwHelpContext = 0;
            strHelpFile = null;

            using (var name = AddressableVariables.CreateBSTR())
            using (var docString = AddressableVariables.CreateBSTR())
            using (var helpContext = AddressableVariables.Create<int>())
            using (var helpFile = AddressableVariables.CreateBSTR())
            {
                int hr = _this_Internal.GetDocumentation(memid, name.Address, docString.Address, helpContext.Address, helpFile.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                strName = name.Value;
                strDocString = docString.Value;
                dwHelpContext = helpContext.Value;
                strHelpFile = helpFile.Value;
            }
        }

        void ComTypes.ITypeInfo.GetDllEntry(int memid, ComTypes.INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
        {
            // for some reason, the ComTypes.ITypeInfo definition for GetDllEntry uses the raw pointers for strings here, 
            // just like our unfriendly version.  This makes it much easier for us to forward on
            int hr = _this_Internal.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
            if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
        }

        void ComTypes.ITypeInfo.GetRefTypeInfo(int hRef, out ComTypes.ITypeInfo ppTI)
        {
            // initialize out parameters
            ppTI = null;

            using (var typeInfoPtr = AddressableVariables.CreateObjectPtr<ComTypes.ITypeInfo>())
            {
                var hr = _this_Internal.GetRefTypeInfo(hRef, typeInfoPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                ppTI = typeInfoPtr.Value;
            }
        }

        void ComTypes.ITypeInfo.AddressOfMember(int memid, ComTypes.INVOKEKIND invKind, out IntPtr ppv)
        {
            // initialize out parameters
            ppv = IntPtr.Zero;

            using (var outPpv = AddressableVariables.Create<IntPtr>())
            {
                int hr = _this_Internal.AddressOfMember(memid, invKind, outPpv.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                ppv = outPpv.Value;
            }
        }

        void ComTypes.ITypeInfo.CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj)
        {
            // initialize out parameters
            ppvObj = null;

            using (var outPpvObj = AddressableVariables.CreateObjectPtr<object>())
            {
                var unkOuter = Marshal.GetIUnknownForObject(pUnkOuter);
                int hr = _this_Internal.CreateInstance(unkOuter, riid, outPpvObj.Address);
                Marshal.Release(unkOuter);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                ppvObj = outPpvObj.Value;
            }
        }

        void ComTypes.ITypeInfo.GetMops(int memid, out string pBstrMops)
        {
            // initialize out parameters
            pBstrMops = null;

            using (var strMops = AddressableVariables.CreateBSTR())
            {
                int hr = _this_Internal.GetMops(memid, strMops.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                pBstrMops = strMops.Value;
            }
        }

        void ComTypes.ITypeInfo.ReleaseTypeAttr(IntPtr pTypeAttr)
        {
            _this_Internal.ReleaseTypeAttr(pTypeAttr);
        }

        void ComTypes.ITypeInfo.ReleaseFuncDesc(IntPtr pFuncDesc)
        {
            _this_Internal.ReleaseFuncDesc(pFuncDesc);
        }

        void ComTypes.ITypeInfo.ReleaseVarDesc(IntPtr pVarDesc)
        {
            _this_Internal.ReleaseVarDesc(pVarDesc);
        }

        // now for the ITypeInfoInternal virtuals to be implemented by the derived class.
        public abstract int GetContainingTypeLib(IntPtr ppTLB, IntPtr pIndex);
        public abstract int GetTypeAttr(IntPtr ppTypeAttr);
        public abstract int GetTypeComp(IntPtr ppTComp);
        public abstract int GetFuncDesc(int index, IntPtr ppFuncDesc);
        public abstract int GetVarDesc(int index, IntPtr ppVarDesc);
        public abstract int GetNames(int memid, IntPtr rgBstrNames, int cMaxNames, IntPtr pcNames);
        public abstract int GetRefTypeOfImplType(int index, IntPtr href);
        public abstract int GetImplTypeFlags(int index, IntPtr pImplTypeFlags);
        public abstract int GetIDsOfNames(IntPtr rgszNames, int cNames, IntPtr pMemId);
        public abstract int Invoke(IntPtr pvInstance, int memid, short wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr);
        public abstract int GetDocumentation(int index, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile);
        public abstract int GetDllEntry(int memid, ComTypes.INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal);
        public abstract int GetRefTypeInfo(int hRef, IntPtr ppTI);
        public abstract int AddressOfMember(int memid, ComTypes.INVOKEKIND invKind, IntPtr ppv);
        public abstract int CreateInstance(IntPtr pUnkOuter, ref Guid riid, IntPtr ppvObj);
        public abstract int GetMops(int memid, IntPtr pBstrMops);
        public abstract void ReleaseTypeAttr(IntPtr pTypeAttr);
        public abstract void ReleaseFuncDesc(IntPtr pFuncDesc);
        public abstract void ReleaseVarDesc(IntPtr pVarDesc);
    }
}
