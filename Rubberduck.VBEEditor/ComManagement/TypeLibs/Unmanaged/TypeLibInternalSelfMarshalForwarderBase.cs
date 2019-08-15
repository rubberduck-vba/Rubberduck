using System;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged
{
    /// <summary>
    /// A compatible version of <see cref="ComTypes.ITypeLib"/>, using <see cref="IntPtr"/> for all out params
    /// see https://msdn.microsoft.com/en-us/library/windows/desktop/ms221549(v=vs.85).aspx
    /// </summary>
    /// <remarks>
    /// We use [PreserveSig] (<see cref="PreserveSigAttribute"/>) so that we can handle HRESULTs directly
    /// </remarks>
    [ComImport(), Guid("00020402-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITypeLibInternal
    {
        [PreserveSig] int GetTypeInfoCount();
        [PreserveSig] /*HRESULT*/ int GetTypeInfo(int index, /*out ITypeInfo*/ IntPtr ppTI);
        [PreserveSig] /*HRESULT*/ int GetTypeInfoType(int index, /*out TYPEKIND*/ IntPtr pTKind);
        [PreserveSig] /*HRESULT*/ int GetTypeInfoOfGuid(ref Guid guid, /*out ITypeInfo*/ IntPtr ppTInfo);
        [PreserveSig] /*HRESULT*/ int GetLibAttr(/*out TLIBATTR*/ IntPtr ppTLibAttr);
        [PreserveSig] /*HRESULT*/ int GetTypeComp(/*out ITypeComp*/ IntPtr ppTComp);
        [PreserveSig] /*HRESULT*/ int GetDocumentation(int index, /*out string*/ IntPtr strName, /*out string*/ IntPtr strDocString, /*out int*/ IntPtr dwHelpContext, /*out string*/ IntPtr strHelpFile);
        [PreserveSig] /*HRESULT*/ int IsName(string szNameBuf, int lHashVal, /*out BOOL*/ IntPtr pfName);
        [PreserveSig] /*HRESULT*/ int FindName(string szNameBuf, int lHashVal, /*out ITypeInfo*/ IntPtr ppTInfo, /*out MEMBERID*/ IntPtr rgMemId, /*out short*/ IntPtr pcFound);
        [PreserveSig] void ReleaseTLibAttr(/*TLIBATTR*/ IntPtr pTLibAttr);
    }

    /// <summary>
    /// This class marshals <see cref="ComTypes.ITypeLib"/> members to <see cref="ITypeLibInternal"/>
    /// </summary>
    /// <remarks>
    /// ITypeLibInternal must be inherited BEFORE ComTypes.ITypeLib as they both have the same IID 
    /// this will ensure QueryInterface(IID_ITypeLib) returns ITypeLibInternal, not ComTypes.ITypeLib
    /// 
    /// These wrappers could likely be implemented much more efficiently if we could use unsafe/fixed code
    /// </remarks>
    internal abstract class TypeLibInternalSelfMarshalForwarderBase : ITypeLibInternal, ComTypes.ITypeLib, IDisposable
    {
        private ITypeLibInternal _this_Internal => (ITypeLibInternal)this;

        private void HandleBadHRESULT(int hr)
        {
            throw RdMarshal.GetExceptionForHR(hr);
        }

        int ComTypes.ITypeLib.GetTypeInfoCount()
        {
            return _this_Internal.GetTypeInfoCount();
        }

        void ComTypes.ITypeLib.GetTypeInfo(int index, out ComTypes.ITypeInfo ppTI)
        {
            // initialize out parameters
            ppTI = default;

            using (var typeInfoPtr = AddressableVariables.CreateObjectPtr<ComTypes.ITypeInfo>())
            {
                var hr = _this_Internal.GetTypeInfo(index, typeInfoPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                ppTI = typeInfoPtr.Value;
            }
        }

        void ComTypes.ITypeLib.GetTypeInfoType(int index, out ComTypes.TYPEKIND pTKind)
        {
            // initialize out parameters
            pTKind = default;

            using (var typeKindPtr = AddressableVariables.Create<ComTypes.TYPEKIND>())
            {
                var hr = _this_Internal.GetTypeInfoType(index, typeKindPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                pTKind = typeKindPtr.Value;
            }
        }

        void ComTypes.ITypeLib.GetTypeInfoOfGuid(ref Guid guid, out ComTypes.ITypeInfo ppTInfo)
        {
            // initialize out parameters
            ppTInfo = default;

            using (var typeInfoPtr = AddressableVariables.CreateObjectPtr<ComTypes.ITypeInfo>())
            {
                var hr = _this_Internal.GetTypeInfoOfGuid(guid, typeInfoPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                ppTInfo = typeInfoPtr.Value;
            }
        }

        void ComTypes.ITypeLib.GetLibAttr(out IntPtr ppTLibAttr)
        {
            // initialize out parameters
            ppTLibAttr = default;

            using (var typeLibAttrPtr = AddressableVariables.CreatePtrTo<ComTypes.TYPELIBATTR>())
            {
                var hr = _this_Internal.GetLibAttr(typeLibAttrPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);
                ppTLibAttr = typeLibAttrPtr.Value.Address;
            }
        }

        void ComTypes.ITypeLib.GetTypeComp(out ComTypes.ITypeComp ppTComp)
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

        void ComTypes.ITypeLib.GetDocumentation(int memid, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile)
        {
            // initialize out parameters
            strName = default;
            strDocString = default;
            dwHelpContext = default;
            strHelpFile = default;

            using (var _name = AddressableVariables.CreateBSTR())
            using (var _docString = AddressableVariables.CreateBSTR())
            using (var _helpContext = AddressableVariables.Create<int>())
            using (var _Helpfile = AddressableVariables.CreateBSTR())
            {
                int hr = _this_Internal.GetDocumentation(memid, _name.Address, _docString.Address, _helpContext.Address, _Helpfile.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                strName = _name.Value;
                strDocString = _docString.Value;
                dwHelpContext = _helpContext.Value;
                strHelpFile = _Helpfile.Value;
            }
        }

        bool ComTypes.ITypeLib.IsName(string szNameBuf, int lHashVal)
        {
            using (var _pfName = AddressableVariables.Create<int>())
            {
                var hr = _this_Internal.IsName(szNameBuf, lHashVal, _pfName.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                return _pfName.Value != 0;
            }
        }

        void ComTypes.ITypeLib.FindName(string szNameBuf, int lHashVal, ComTypes.ITypeInfo[] ppTInfo, int[] rgMemId, ref short pcFound)
        {
            // We can't use the managed arrays as passed in.  We create our own unmanaged arrays, 
            // and copy them into the managed ones on completion
            using (var _ppTInfo = AddressableVariables.CreateObjectPtr<ComTypes.ITypeInfo>(pcFound))
            using (var _MemIds = AddressableVariables.Create<int>(pcFound))
            using (var _pcFound = AddressableVariables.Create<short>())
            {
                var hr = _this_Internal.FindName(szNameBuf, lHashVal, _ppTInfo.Address, _MemIds.Address, _pcFound.Address);
                if (ComHelper.HRESULT_FAILED(hr)) HandleBadHRESULT(hr);

                _ppTInfo.CopyArrayTo(ppTInfo);
                _MemIds.CopyArrayTo(rgMemId);
                pcFound = _pcFound.Value;
            }
        }

        void ComTypes.ITypeLib.ReleaseTLibAttr(IntPtr pTLibAttr)
        {
            _this_Internal.ReleaseTLibAttr(pTLibAttr);
        }

        // now for the ITypeLibInternal virtuals to be implemented by the derived class.
        public abstract int GetTypeInfoCount();
        public abstract int GetTypeInfo(int index, IntPtr ppTI);
        public abstract int GetTypeInfoType(int index, IntPtr pTKind);
        public abstract int GetTypeInfoOfGuid(ref Guid guid, IntPtr ppTInfo);
        public abstract int GetLibAttr(IntPtr ppTLibAttr);
        public abstract int GetTypeComp(IntPtr ppTComp);
        public abstract int GetDocumentation(int index, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile);
        public abstract int IsName(string szNameBuf, int lHashVal, IntPtr pfName);
        public abstract int FindName(string szNameBuf, int lHashVal, IntPtr ppTInfo, IntPtr rgMemId, IntPtr pcFound);
        public abstract void ReleaseTLibAttr(IntPtr pTLibAttr);

        public abstract void Dispose();
    }
}
