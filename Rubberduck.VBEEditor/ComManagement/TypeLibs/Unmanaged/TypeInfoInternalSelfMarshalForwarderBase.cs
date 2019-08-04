using System;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged
{
    /// <summary>
    /// This class marshals <see cref="ComTypes.ITypeInfo"/> members to <see cref="ITypeInfoInternal"/>
    /// </summary>
    /// <remarks>
    /// ITypeInfoInternal must be inherited BEFORE ComTypes.ITypeInfo as they both have the same IID 
    /// this will ensure QueryInterface(IID_ITypeInfo) returns ITypeInfoInternal, not ComTypes.ITypeInfo
    /// </remarks>
    internal abstract class TypeInfoInternalSelfMarshalForwarderBase : ITypeInfoInternal, ComTypes.ITypeInfo
    {
        private ITypeInfoInternal _this_Internal => (ITypeInfoInternal)this;

        private static void HandleBadHRESULT(int hr)
        {
            throw RdMarshal.GetExceptionForHR(hr);
        }

        void ComTypes.ITypeInfo.GetContainingTypeLib(out ComTypes.ITypeLib ppTLB, out int pIndex)
        {
            // initialize out parameters
            ppTLB = default;
            pIndex = default;

            using (var typeLibPtr = AddressableVariables.CreateObjectPtr<ComTypes.ITypeLib>())
            using (var indexPtr = AddressableVariables.Create<int>())
            {
                var hr = _this_Internal.GetContainingTypeLib(typeLibPtr.Address, indexPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                ppTLB = typeLibPtr.Value;
                pIndex = indexPtr.Value;
            }
        }

        void ComTypes.ITypeInfo.GetTypeAttr(out IntPtr ppTypeAttr)
        {
            // initialize out parameters
            ppTypeAttr = default;

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
            ppTComp = default;

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
            ppFuncDesc = default;

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
            ppVarDesc = default;

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
            pcNames = default;

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
            href = default;

            using (var outHref = AddressableVariables.Create<int>())
            {
                var hr = _this_Internal.GetRefTypeOfImplType(index, outHref.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                href = outHref.Value;
            }
        }

        void ComTypes.ITypeInfo.GetImplTypeFlags(int index, out ComTypes.IMPLTYPEFLAGS pImplTypeFlags)
        {
            // initialize out parameters
            pImplTypeFlags = default;

            using (var implTypeFlags = AddressableVariables.Create<ComTypes.IMPLTYPEFLAGS>())
            {
                var hr = _this_Internal.GetImplTypeFlags(index, implTypeFlags.Address);
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
            strName = default;
            strDocString = default;
            dwHelpContext = default;
            strHelpFile = default;

            using (var name = AddressableVariables.CreateBSTR())
            using (var docString = AddressableVariables.CreateBSTR())
            using (var helpContext = AddressableVariables.Create<int>())
            using (var helpFile = AddressableVariables.CreateBSTR())
            {
                var hr = _this_Internal.GetDocumentation(memid, name.Address, docString.Address, helpContext.Address, helpFile.Address);
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
            var hr = _this_Internal.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
            if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
        }

        void ComTypes.ITypeInfo.GetRefTypeInfo(int hRef, out ComTypes.ITypeInfo ppTI)
        {
            // initialize out parameters
            ppTI = default;

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
            ppv = default;

            using (var outPpv = AddressableVariables.Create<IntPtr>())
            {
                var hr = _this_Internal.AddressOfMember(memid, invKind, outPpv.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                ppv = outPpv.Value;
            }
        }

        void ComTypes.ITypeInfo.CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj)
        {
            // initialize out parameters
            ppvObj = default;

            using (var outPpvObj = AddressableVariables.CreateObjectPtr<object>())
            {
                var unkOuter = RdMarshal.GetIUnknownForObject(pUnkOuter);
                var hr = _this_Internal.CreateInstance(unkOuter, riid, outPpvObj.Address);
                RdMarshal.Release(unkOuter);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                ppvObj = outPpvObj.Value;
            }
        }

        void ComTypes.ITypeInfo.GetMops(int memid, out string pBstrMops)
        {
            // initialize out parameters
            pBstrMops = default;

            using (var strMops = AddressableVariables.CreateBSTR())
            {
                var hr = _this_Internal.GetMops(memid, strMops.Address);
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
