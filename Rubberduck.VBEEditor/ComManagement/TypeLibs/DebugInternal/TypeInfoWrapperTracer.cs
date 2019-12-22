using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.DebugInternal
{
    /// <summary>
    /// Wraps the existing implementation so that we can trace the calls for
    /// diagnostics or debugging. See <see cref="TypeApiFactory"/> for
    /// creating a class to be traced. The class should not be created directly.
    /// </summary>
    internal class TypeInfoWrapperTracer : ITypeInfoWrapper, ITypeInfoInternal
    {
        private readonly ITypeInfoWrapper _wrapper;
        private readonly ITypeInfoInternal _inner;

        internal TypeInfoWrapperTracer(ITypeInfoWrapper wrapper, ITypeInfoInternal inner)
        {
            _wrapper = wrapper;
            _inner = inner;
        }

        private static void Before(string parameters = null, [CallerMemberName] string methodName = null)
        {
            System.Diagnostics.Debug.Print($"Entering {nameof(ITypeInfoWrapper)}::{methodName}; {parameters}");
        }

        private static void After(string parameters = null, [CallerMemberName] string methodName = null)
        {
            System.Diagnostics.Debug.Print($"Leaving {nameof(ITypeInfoWrapper)}::{methodName}; {parameters}");
        }

        void ITypeInfo.GetTypeAttr(out IntPtr ppTypeAttr)
        {
            Before();
            _wrapper.GetTypeAttr(out var t);
            After($"{nameof(ppTypeAttr)}: {t}");
            ppTypeAttr = t;
        }

        void ITypeInfo.GetTypeComp(out ITypeComp ppTComp)
        {
            Before();
            _wrapper.GetTypeComp(out var t);
            After($"{nameof(ppTComp)}: {t?.GetHashCode()}");
            ppTComp = t;
        }

        void ITypeInfo.GetFuncDesc(int index, out IntPtr ppFuncDesc)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetFuncDesc(index, out var t);
            After($"{nameof(ppFuncDesc)}: {t}");
            ppFuncDesc = t;
        }

        void ITypeInfo.GetVarDesc(int index, out IntPtr ppVarDesc)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetVarDesc(index, out var t);
            After($"{nameof(ppVarDesc)}: {t}");
            ppVarDesc = t;
        }

        void ITypeInfo.GetNames(int memid, string[] rgBstrNames, int cMaxNames, out int pcNames)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(rgBstrNames)}: {(rgBstrNames == null ? "null" : "strings")}, {nameof(cMaxNames)}: {cMaxNames}");
            _wrapper.GetNames(memid, rgBstrNames, cMaxNames, out var t);
            After($"{nameof(rgBstrNames)}: {(rgBstrNames == null ? "null" : "strings")}, {nameof(pcNames)}: {t}");
            pcNames = t;
        }

        void ITypeInfo.GetRefTypeOfImplType(int index, out int href)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetRefTypeOfImplType(index, out var t);
            After($"{nameof(href)}: {t}");
            href = t;
        }

        void ITypeInfo.GetImplTypeFlags(int index, out IMPLTYPEFLAGS pImplTypeFlags)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetImplTypeFlags(index, out var t);
            After($"{nameof(pImplTypeFlags)}: {t}");
            pImplTypeFlags = t;
        }

        void ITypeInfo.GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId)
        {
            Before($"{nameof(rgszNames)}: {(rgszNames == null ? "null" : "strings")}, {nameof(cNames)}: {cNames}, {nameof(pMemId)}: {(pMemId == null ? "null" : "ints")}");
            _wrapper.GetIDsOfNames(rgszNames, cNames, pMemId);
            After($"{nameof(rgszNames)}: {(rgszNames == null ? "null" : "strings")}, {nameof(pMemId)}: {(pMemId == null ? "null" : "ints")}");
        }

        void ITypeInfo.Invoke(object pvInstance, int memid, short wFlags, ref DISPPARAMS pDispParams, IntPtr pVarResult,
            IntPtr pExcepInfo, out int puArgErr)
        {
            Before("parameters not included");
            _wrapper.Invoke(pvInstance, memid, wFlags, ref pDispParams, pVarResult, pExcepInfo, out puArgErr);
            After("parameters not included");
        }

        public void GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext,
            out string strHelpFile)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetDocumentation(index, out var t1, out var t2, out var t3, out var t4);
            After($"{nameof(strName)}: {t1}, {nameof(strDocString)}: {t2}, {nameof(dwHelpContext)}: {t3}, {nameof(strHelpFile)}: {t4}");
            strName = t1;
            strDocString = t2;
            dwHelpContext = t3;
            strHelpFile = t4;
        }

        int ITypeInfoWrapper.GetDllEntry(int memid, INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(invKind)}: {invKind}, {nameof(pBstrDllName)}: {pBstrDllName}, {nameof(pBstrName)}: {pBstrName}, {nameof(pwOrdinal)}: {pwOrdinal}");
            var result = _wrapper.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetDllEntry(int memid, INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(invKind)}: {invKind}, {nameof(pBstrDllName)}: {pBstrDllName}, {nameof(pBstrName)}: {pBstrName}, {nameof(pwOrdinal)}: {pwOrdinal}");
            var result = _inner.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
            After($"{nameof(result)}: {result}, {nameof(pBstrDllName)}: {pBstrDllName}, {nameof(pBstrName)}: {pBstrName}, {nameof(pwOrdinal)}: {pwOrdinal}");
            return result;
        }

        private int GetRefTypeInfo(int hRef, IntPtr ppTI)
        {
            Before($"{nameof(hRef)}: {hRef}, {nameof(ppTI)}: {ppTI}");
            var result = _wrapper.GetRefTypeInfo(hRef, ppTI);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetRefTypeInfo(int hRef, IntPtr ppTI)
        {
            return GetRefTypeInfo(hRef, ppTI);
        }

        int ITypeInfoWrapper.GetRefTypeInfo(int hRef, IntPtr ppTI)
        {
            return GetRefTypeInfo(hRef, ppTI);
        }

        private int AddressOfMember(int memid, INVOKEKIND invKind, IntPtr ppv)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(invKind)}: {invKind}, {nameof(ppv)}: {ppv}");
            var result = _wrapper.AddressOfMember(memid, invKind, ppv);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.AddressOfMember(int memid, INVOKEKIND invKind, IntPtr ppv)
        {
            return AddressOfMember(memid, invKind, ppv);
        }

        int ITypeInfoWrapper.AddressOfMember(int memid, INVOKEKIND invKind, IntPtr ppv)
        {
            return AddressOfMember(memid, invKind, ppv);
        }

        private int CreateInstance(IntPtr pUnkOuter, ref Guid riid, IntPtr ppvObj)
        {
            Before($"{nameof(pUnkOuter)}: {pUnkOuter}, {nameof(riid)}: {riid}, {nameof(ppvObj)}: {ppvObj}");
            var result = _wrapper.CreateInstance(pUnkOuter, ref riid, ppvObj);
            After($"{nameof(result)}: {result}, {nameof(riid)}: {riid}");
            return result;
        }

        int ITypeInfoInternal.CreateInstance(IntPtr pUnkOuter, ref Guid riid, IntPtr ppvObj)
        {
            return CreateInstance(pUnkOuter, ref riid, ppvObj);
        }

        int ITypeInfoWrapper.CreateInstance(IntPtr pUnkOuter, ref Guid riid, IntPtr ppvObj)
        {
            return CreateInstance(pUnkOuter, ref riid, ppvObj);
        }

        private int GetMops(int memid, IntPtr pBstrMops)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(pBstrMops)}: {pBstrMops}");
            var result = _wrapper.GetMops(memid, pBstrMops);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetMops(int memid, IntPtr pBstrMops)
        {
            return GetMops(memid, pBstrMops);
        }

        int ITypeInfoWrapper.GetMops(int memid, IntPtr pBstrMops)
        {
            return GetMops(memid, pBstrMops);
        }

        private void ReleaseTypeAttr(IntPtr pTypeAttr)
        {
            Before($"{nameof(pTypeAttr)}: {pTypeAttr}");
            _wrapper.ReleaseTypeAttr(pTypeAttr);
            After();
        }

        void ITypeInfoInternal.ReleaseTypeAttr(IntPtr pTypeAttr)
        {
            ReleaseTypeAttr(pTypeAttr);
        }

        void ITypeInfo.ReleaseTypeAttr(IntPtr pTypeAttr)
        {
            ReleaseTypeAttr(pTypeAttr);
        }

        void ITypeInfoWrapper.ReleaseTypeAttr(IntPtr pTypeAttr)
        {
            ReleaseTypeAttr(pTypeAttr);
        }

        public void ReleaseFuncDesc(IntPtr pFuncDesc)
        {
            Before($"{nameof(pFuncDesc)}: {pFuncDesc}");
            _wrapper.ReleaseFuncDesc(pFuncDesc);
            After();
        }

        private void ReleaseVarDesc(IntPtr pVarDesc)
        {
            Before($"{nameof(pVarDesc)}: {pVarDesc}");
            _wrapper.ReleaseVarDesc(pVarDesc);  
            After();
        }

        void ITypeInfoInternal.ReleaseVarDesc(IntPtr pVarDesc)
        {
            ReleaseVarDesc(pVarDesc);
        }

        void ITypeInfo.ReleaseVarDesc(IntPtr pVarDesc)
        {
            ReleaseVarDesc(pVarDesc);
        }

        void ITypeInfoWrapper.ReleaseVarDesc(IntPtr pVarDesc)
        {
            ReleaseVarDesc(pVarDesc);
        }

        ITypeInfoFunctionCollection ITypeInfoWrapper.Funcs
        {
            get
            {
                Before();
                var result = _wrapper.Funcs;
                After($"{result}: {result}");
                return result;
            }
        }

        ITypeInfoVariablesCollection ITypeInfoWrapper.Vars
        {
            get
            {
                Before();
                var result = _wrapper.Vars;
                After($"{result}: {result}");
                return result;
            }
        }

        ITypeInfoImplementedInterfacesCollection ITypeInfoWrapper.ImplementedInterfaces
        {
            get
            {
                Before();
                var result = _wrapper.ImplementedInterfaces;
                After($"{result}: {result}");
                return result;
            }
        }

        private int GetDocumentation(int memid, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(strName)}: {strName}, {nameof(strDocString)}: {strDocString}, {nameof(dwHelpContext)}: {dwHelpContext}, {nameof(strHelpFile)}: {strHelpFile}");
            var result = _wrapper.GetDocumentation(memid, strName, strDocString, dwHelpContext, strHelpFile);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetDocumentation(int memid, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            return GetDocumentation(memid, strName, strDocString, dwHelpContext, strHelpFile);
        }

        int ITypeInfoWrapper.GetDocumentation(int memid, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            return GetDocumentation(memid, strName, strDocString, dwHelpContext, strHelpFile);
        }

        void ITypeInfo.GetDllEntry(int memid, INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(invKind)}: {invKind}, {nameof(pBstrDllName)}: {pBstrDllName}, {nameof(pBstrName)}: {pBstrName}, {nameof(pwOrdinal)}: {pwOrdinal}");
            var result = _inner.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
            After($"{nameof(result)}: {result}");
        }

        void ITypeInfo.GetRefTypeInfo(int hRef, out ITypeInfo ppTI)
        {
            Before($"{nameof(hRef)}: {hRef}");
            _wrapper.GetRefTypeInfo(hRef, out var t);
            After($"{nameof(ppTI)}: {t?.GetHashCode()}");
            ppTI = t;
        }

        void ITypeInfo.AddressOfMember(int memid, INVOKEKIND invKind, out IntPtr ppv)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(invKind)}: {invKind}");
            _wrapper.AddressOfMember(memid, invKind, out var t);
            After($"{nameof(ppv)}: {t}");
            ppv = t;
        }

        void ITypeInfo.CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj)
        {
            Before($"{nameof(pUnkOuter)}: {pUnkOuter?.GetHashCode()}, {nameof(riid)}: {riid}");
            _wrapper.CreateInstance(pUnkOuter, ref riid, out var t);
            After($"{nameof(riid)}: {riid}, {nameof(ppvObj)}: {t?.GetHashCode()}");
            ppvObj = t;
        }

        void ITypeInfo.GetMops(int memid, out string pBstrMops)
        {
            Before($"{nameof(memid)}: {memid}");
            _wrapper.GetMops(memid, out var t);
            After($"{nameof(pBstrMops)}: {t}");
            pBstrMops = t;
        }

        void ITypeInfo.GetContainingTypeLib(out ITypeLib ppTLB, out int pIndex)
        {
            Before();
            _wrapper.GetContainingTypeLib(out var t1, out var t2);
            After($"{nameof(ppTLB)}: {t1?.GetHashCode()}, {nameof(pIndex)}: {t2}");
            ppTLB = t1;
            pIndex = t2;
        }

        ITypeLib ITypeInfoWrapper.Container
        {
            get
            {
                Before();
                var result = _wrapper.Container;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        int ITypeInfoWrapper.ContainerIndex
        {
            get
            {
                Before();
                var result = _wrapper.ContainerIndex;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        bool ITypeInfoWrapper.HasModuleScopeCompilationErrors
        {
            get
            {
                Before();
                var result = _wrapper.HasModuleScopeCompilationErrors;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        bool ITypeInfoWrapper.HasVBEExtensions
        {
            get
            {
                Before();
                var result = _wrapper.HasVBEExtensions;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        TYPEATTR ITypeInfoWrapper.CachedAttributes
        {
            get
            {
                Before();
                var result = _wrapper.CachedAttributes;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        bool ITypeInfoWrapper.HasSimulatedContainer
        {
            get
            {
                Before();
                var result = _wrapper.HasSimulatedContainer;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        bool ITypeInfoWrapper.IsUserFormBaseClass
        {
            get
            {
                Before();
                var result = _wrapper.IsUserFormBaseClass;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        string ITypeInfoWrapper.Name
        {
            get
            {
                Before();
                var result = _wrapper.Name;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        string ITypeInfoWrapper.DocString
        {
            get
            {
                Before();
                var result = _wrapper.DocString;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        int ITypeInfoWrapper.HelpContext
        {
            get
            {
                Before();
                var result = _wrapper.HelpContext;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        string ITypeInfoWrapper.HelpFile
        {
            get
            {
                Before();
                var result = _wrapper.HelpFile;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        string ITypeInfoWrapper.ProgID
        {
            get
            {
                Before();
                var result = _wrapper.ProgID;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        Guid ITypeInfoWrapper.GUID
        {
            get
            {
                Before();
                var result = _wrapper.GUID;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        TYPEKIND_VBE ITypeInfoWrapper.TypeKind
        {
            get
            {
                Before();
                var result = _wrapper.TypeKind;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        bool ITypeInfoWrapper.HasPredeclaredId
        {
            get
            {
                Before();
                var result = _wrapper.HasPredeclaredId;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        TYPEFLAGS ITypeInfoWrapper.Flags
        {
            get
            {
                Before();
                var result = _wrapper.Flags;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        string ITypeInfoWrapper.ContainerName
        {
            get
            {
                Before();
                var result = _wrapper.ContainerName;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        ITypeInfoVBEExtensions ITypeInfoWrapper.VBEExtensions
        {
            get
            {
                Before();
                var result = _wrapper.VBEExtensions;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public void Dispose()
        {
            Before();
            _wrapper.Dispose();
            After();
        }

        int ITypeInfoWrapper.GetSafeRefTypeInfo(int hRef, out ITypeInfoWrapper outTI)
        {
            Before($"{nameof(hRef)}: {hRef}");
            var result = _wrapper.GetSafeRefTypeInfo(hRef, out var t);
            After($"{nameof(result)}: {result}, {nameof(outTI)}: {t?.GetHashCode()}");
            outTI = t;
            return result;
        }

        IntPtr ITypeInfoWrapper.GetCOMReferencePtr()
        {
            Before();
            var result = _wrapper.GetCOMReferencePtr();
            After($"{nameof(result)}: {result}");
            return result;
        }

        private int GetContainingTypeLib(IntPtr ppTLB, IntPtr pIndex)
        {
            Before($"{nameof(ppTLB)}: {ppTLB}, {nameof(pIndex)}: {pIndex}");
            var result = _wrapper.GetContainingTypeLib(ppTLB, pIndex);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetContainingTypeLib(IntPtr ppTLB, IntPtr pIndex)
        {
            return GetContainingTypeLib(ppTLB, pIndex);
        }

        int ITypeInfoWrapper.GetContainingTypeLib(IntPtr ppTLB, IntPtr pIndex)
        {
            return GetContainingTypeLib(ppTLB, pIndex);
        }

        private int GetTypeAttr(IntPtr ppTypeAttr)
        {
            Before($"{nameof(ppTypeAttr)}: {ppTypeAttr}");
            var result = _wrapper.GetTypeAttr(ppTypeAttr);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetTypeAttr(IntPtr ppTypeAttr)
        {
            return GetTypeAttr(ppTypeAttr);
        }

        int ITypeInfoWrapper.GetTypeAttr(IntPtr ppTypeAttr)
        {
            return GetTypeAttr(ppTypeAttr);
        }

        private int GetTypeComp(IntPtr ppTComp)
        {
            Before($"{nameof(ppTComp)}: {ppTComp}");
            var result = _wrapper.GetTypeComp(ppTComp);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetTypeComp(IntPtr ppTComp)
        {
            return GetTypeComp(ppTComp);
        }

        int ITypeInfoWrapper.GetTypeComp(IntPtr ppTComp)
        {
            return GetTypeComp(ppTComp);
        }

        private int GetFuncDesc(int index, IntPtr ppFuncDesc)
        {
            Before($"{nameof(index)}: {index}, {nameof(ppFuncDesc)}: {ppFuncDesc}");
            var result = _wrapper.GetFuncDesc(index, ppFuncDesc);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetFuncDesc(int index, IntPtr ppFuncDesc)
        {
            return GetFuncDesc(index, ppFuncDesc);
        }

        int ITypeInfoWrapper.GetFuncDesc(int index, IntPtr ppFuncDesc)
        {
            return GetFuncDesc(index, ppFuncDesc);
        }

        public int GetVarDesc(int index, IntPtr ppVarDesc)
        {
            Before($"{nameof(index)}: {index}, {nameof(ppVarDesc)}: {ppVarDesc}");
            var result = _wrapper.GetVarDesc(index, ppVarDesc);
            After($"{nameof(result)}: {result}");
            return result;
        }

        private int GetNames(int memid, IntPtr rgBstrNames, int cMaxNames, IntPtr pcNames)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(rgBstrNames)}: {rgBstrNames}, {nameof(cMaxNames)}: {cMaxNames}: {nameof(pcNames)}: {pcNames}");
            var result = _wrapper.GetNames(memid, rgBstrNames, cMaxNames, pcNames);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetNames(int memid, IntPtr rgBstrNames, int cMaxNames, IntPtr pcNames)
        {
            return GetNames(memid, rgBstrNames, cMaxNames, pcNames);
        }

        int ITypeInfoWrapper.GetNames(int memid, IntPtr rgBstrNames, int cMaxNames, IntPtr pcNames)
        {
            return GetNames(memid, rgBstrNames, cMaxNames, pcNames);
        }

        private int GetRefTypeOfImplType(int index, IntPtr href)
        {
            Before($"{nameof(index)}: {index}, {nameof(href)}: {href}");
            var result = _wrapper.GetRefTypeOfImplType(index, href);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetRefTypeOfImplType(int index, IntPtr href)
        {
            return GetRefTypeOfImplType(index, href);
        }

        int ITypeInfoWrapper.GetRefTypeOfImplType(int index, IntPtr href)
        {
            return GetRefTypeOfImplType(index, href);
        }

        private int GetImplTypeFlags(int index, IntPtr pImplTypeFlags)
        {
            Before($"{nameof(index)}: {index}, {nameof(pImplTypeFlags)}: {pImplTypeFlags}");
            var result = _wrapper.GetImplTypeFlags(index, pImplTypeFlags);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetImplTypeFlags(int index, IntPtr pImplTypeFlags)
        {
            return GetImplTypeFlags(index, pImplTypeFlags);
        }

        int ITypeInfoWrapper.GetImplTypeFlags(int index, IntPtr pImplTypeFlags)
        {
            return GetImplTypeFlags(index, pImplTypeFlags);
        }

        private int GetIDsOfNames(IntPtr rgszNames, int cNames, IntPtr pMemId)
        {
            Before($"{nameof(rgszNames)}: {rgszNames}, {nameof(cNames)}: {cNames}, {nameof(pMemId)}: {pMemId}");
            var result = _wrapper.GetIDsOfNames(rgszNames, cNames, pMemId);
            After($"{nameof(result)}: {result}");
            return result;
        }

        int ITypeInfoInternal.GetIDsOfNames(IntPtr rgszNames, int cNames, IntPtr pMemId)
        {
            return GetIDsOfNames(rgszNames, cNames, pMemId);
        }

        int ITypeInfoWrapper.GetIDsOfNames(IntPtr rgszNames, int cNames, IntPtr pMemId)
        {
            return GetIDsOfNames(rgszNames, cNames, pMemId);
        }

        private int Invoke(IntPtr pvInstance, int memid, short wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo,
            IntPtr puArgErr)
        {
            Before($"parameters not supplied");
            var result = _wrapper.Invoke(pvInstance, memid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
            After($"parameters not supplied");
            return result;
        }

        int ITypeInfoInternal.Invoke(IntPtr pvInstance, int memid, short wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo,
            IntPtr puArgErr)
        {
            return Invoke(pvInstance, memid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
        }

        int ITypeInfoWrapper.Invoke(IntPtr pvInstance, int memid, short wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo,
            IntPtr puArgErr)
        {
            return Invoke(pvInstance, memid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
        }
    }
}