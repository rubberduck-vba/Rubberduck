using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.DebugInternal
{
    internal class TypeInfoWrapperTracer : ITypeInfoWrapper
    {
        private readonly ITypeInfoWrapper _wrapper;

        internal TypeInfoWrapperTracer(ITypeInfoWrapper wrapper)
        {
            _wrapper = wrapper;
        }

        private static void Before(string parameters = null, [CallerMemberName] string methodName = null)
        {
            System.Diagnostics.Debug.Print($"Entering {nameof(ITypeInfoWrapper)}::{methodName}; {parameters}");
        }

        private static void After(string parameters = null, [CallerMemberName] string methodName = null)
        {
            System.Diagnostics.Debug.Print($"Leaving {nameof(ITypeInfoWrapper)}::{methodName}; {parameters}");
        }

        public void GetTypeAttr(out IntPtr ppTypeAttr)
        {
            Before();
            _wrapper.GetTypeAttr(out var t);
            After($"{nameof(ppTypeAttr)}: {t}");
            ppTypeAttr = t;
        }

        public void GetTypeComp(out ITypeComp ppTComp)
        {
            Before();
            _wrapper.GetTypeComp(out var t);
            After($"{nameof(ppTComp)}: {t?.GetHashCode()}");
            ppTComp = t;
        }

        public void GetFuncDesc(int index, out IntPtr ppFuncDesc)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetFuncDesc(index, out var t);
            After($"{nameof(ppFuncDesc)}: {t}");
            ppFuncDesc = t;
        }

        public void GetVarDesc(int index, out IntPtr ppVarDesc)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetVarDesc(index, out var t);
            After($"{nameof(ppVarDesc)}: {t}");
            ppVarDesc = t;
        }

        public void GetNames(int memid, string[] rgBstrNames, int cMaxNames, out int pcNames)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(rgBstrNames)}: {(rgBstrNames == null ? "null" : "strings")}, {nameof(cMaxNames)}: {cMaxNames}");
            _wrapper.GetNames(memid, rgBstrNames, cMaxNames, out var t);
            After($"{nameof(rgBstrNames)}: {(rgBstrNames == null ? "null" : "strings")}, {nameof(pcNames)}: {t}");
            pcNames = t;
        }

        public void GetRefTypeOfImplType(int index, out int href)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetRefTypeOfImplType(index, out var t);
            After($"{nameof(href)}: {t}");
            href = t;
        }

        public void GetImplTypeFlags(int index, out IMPLTYPEFLAGS pImplTypeFlags)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetImplTypeFlags(index, out var t);
            After($"{nameof(pImplTypeFlags)}: {t}");
            pImplTypeFlags = t;
        }

        public void GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId)
        {
            Before($"{nameof(rgszNames)}: {(rgszNames == null ? "null" : "strings")}, {nameof(cNames)}: {cNames}, {nameof(pMemId)}: {(pMemId == null ? "null" : "ints")}");
            _wrapper.GetIDsOfNames(rgszNames, cNames, pMemId);
            After($"{nameof(rgszNames)}: {(rgszNames == null ? "null" : "strings")}, {nameof(pMemId)}: {(pMemId == null ? "null" : "ints")}");
        }

        public void Invoke(object pvInstance, int memid, short wFlags, ref DISPPARAMS pDispParams, IntPtr pVarResult,
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

        public int GetRefTypeInfo(int hRef, IntPtr ppTI)
        {
            Before($"{nameof(hRef)}: {hRef}, {nameof(ppTI)}: {ppTI}");
            var result = _wrapper.GetRefTypeInfo(hRef, ppTI);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int AddressOfMember(int memid, INVOKEKIND invKind, IntPtr ppv)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(invKind)}: {invKind}, {nameof(ppv)}: {ppv}");
            var result = _wrapper.AddressOfMember(memid, invKind, ppv);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int CreateInstance(IntPtr pUnkOuter, ref Guid riid, IntPtr ppvObj)
        {
            Before($"{nameof(pUnkOuter)}: {pUnkOuter}, {nameof(riid)}: {riid}, {nameof(ppvObj)}: {ppvObj}");
            var result = _wrapper.CreateInstance(pUnkOuter, ref riid, ppvObj);
            After($"{nameof(result)}: {result}, {nameof(riid)}: {riid}");
            return result;
        }

        public int GetMops(int memid, IntPtr pBstrMops)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(pBstrMops)}: {pBstrMops}");
            var result = _wrapper.GetMops(memid, pBstrMops);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public void ReleaseTypeAttr(IntPtr pTypeAttr)
        {
            Before($"{nameof(pTypeAttr)}: {pTypeAttr}");
            _wrapper.ReleaseTypeAttr(pTypeAttr);
            After();
        }
        
        public void ReleaseFuncDesc(IntPtr pFuncDesc)
        {
            Before($"{nameof(pFuncDesc)}: {pFuncDesc}");
            _wrapper.ReleaseFuncDesc(pFuncDesc);
            After();
        }

        public void ReleaseVarDesc(IntPtr pVarDesc)
        {
            Before($"{nameof(pVarDesc)}: {pVarDesc}");
            _wrapper.ReleaseVarDesc(pVarDesc);  
            After();
        }

        public ITypeInfoFunctionCollection Funcs
        {
            get
            {
                Before();
                var result = _wrapper.Funcs;
                After($"{result}: {result}");
                return result;
            }
        }

        public ITypeInfoVariablesCollection Vars
        {
            get
            {
                Before();
                var result = _wrapper.Vars;
                After($"{result}: {result}");
                return result;
            }
        }

        public ITypeInfoImplementedInterfacesCollection ImplementedInterfaces
        {
            get
            {
                Before();
                var result = _wrapper.ImplementedInterfaces;
                After($"{result}: {result}");
                return result;
            }
        }

        public int GetDocumentation(int memid, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(strName)}: {strName}, {nameof(strDocString)}: {strDocString}, {nameof(dwHelpContext)}: {dwHelpContext}, {nameof(strHelpFile)}: {strHelpFile}");
            var result = _wrapper.GetDocumentation(memid, strName, strDocString, dwHelpContext, strHelpFile);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public void GetDllEntry(int memid, INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(invKind)}: {invKind}, {nameof(pBstrDllName)}: {pBstrDllName}, {nameof(pBstrName)}: {pBstrName}, {nameof(pwOrdinal)}: {pwOrdinal}");
            var result = _wrapper.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
            After($"{nameof(result)}: {result}");
        }

        public void GetRefTypeInfo(int hRef, out ITypeInfo ppTI)
        {
            Before($"{nameof(hRef)}: {hRef}");
            _wrapper.GetRefTypeInfo(hRef, out var t);
            After($"{nameof(ppTI)}: {t?.GetHashCode()}");
            ppTI = t;
        }

        public void AddressOfMember(int memid, INVOKEKIND invKind, out IntPtr ppv)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(invKind)}: {invKind}");
            _wrapper.AddressOfMember(memid, invKind, out var t);
            After($"{nameof(ppv)}: {t}");
            ppv = t;
        }

        public void CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj)
        {
            Before($"{nameof(pUnkOuter)}: {pUnkOuter?.GetHashCode()}, {nameof(riid)}: {riid}");
            _wrapper.CreateInstance(pUnkOuter, ref riid, out var t);
            After($"{nameof(riid)}: {riid}, {nameof(ppvObj)}: {t?.GetHashCode()}");
            ppvObj = t;
        }

        public void GetMops(int memid, out string pBstrMops)
        {
            Before($"{nameof(memid)}: {memid}");
            _wrapper.GetMops(memid, out var t);
            After($"{nameof(pBstrMops)}: {t}");
            pBstrMops = t;
        }

        public void GetContainingTypeLib(out ITypeLib ppTLB, out int pIndex)
        {
            Before();
            _wrapper.GetContainingTypeLib(out var t1, out var t2);
            After($"{nameof(ppTLB)}: {t1?.GetHashCode()}, {nameof(pIndex)}: {t2}");
            ppTLB = t1;
            pIndex = t2;
        }
        
        public ITypeLib Container
        {
            get
            {
                Before();
                var result = _wrapper.Container;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public int ContainerIndex
        {
            get
            {
                Before();
                var result = _wrapper.ContainerIndex;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public bool HasModuleScopeCompilationErrors
        {
            get
            {
                Before();
                var result = _wrapper.HasModuleScopeCompilationErrors;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public bool HasVBEExtensions
        {
            get
            {
                Before();
                var result = _wrapper.HasVBEExtensions;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public TYPEATTR CachedAttributes
        {
            get
            {
                Before();
                var result = _wrapper.CachedAttributes;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public bool HasSimulatedContainer
        {
            get
            {
                Before();
                var result = _wrapper.HasSimulatedContainer;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public bool IsUserFormBaseClass
        {
            get
            {
                Before();
                var result = _wrapper.IsUserFormBaseClass;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public string Name
        {
            get
            {
                Before();
                var result = _wrapper.Name;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public string DocString
        {
            get
            {
                Before();
                var result = _wrapper.DocString;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public int HelpContext
        {
            get
            {
                Before();
                var result = _wrapper.HelpContext;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public string HelpFile
        {
            get
            {
                Before();
                var result = _wrapper.HelpFile;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public string ProgID
        {
            get
            {
                Before();
                var result = _wrapper.ProgID;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public Guid GUID
        {
            get
            {
                Before();
                var result = _wrapper.GUID;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public TYPEKIND_VBE TypeKind
        {
            get
            {
                Before();
                var result = _wrapper.TypeKind;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public bool HasPredeclaredId
        {
            get
            {
                Before();
                var result = _wrapper.HasPredeclaredId;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public TYPEFLAGS Flags
        {
            get
            {
                Before();
                var result = _wrapper.Flags;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public string ContainerName
        {
            get
            {
                Before();
                var result = _wrapper.ContainerName;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public ITypeInfoVBEExtensions VBEExtensions
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

        public int GetSafeRefTypeInfo(int hRef, out ITypeInfoWrapper outTI)
        {
            Before($"{nameof(hRef)}: {hRef}");
            var result = _wrapper.GetSafeRefTypeInfo(hRef, out var t);
            After($"{nameof(result)}: {result}, {nameof(outTI)}: {t?.GetHashCode()}");
            outTI = t;
            return result;
        }

        public IntPtr GetCOMReferencePtr()
        {
            Before();
            var result = _wrapper.GetCOMReferencePtr();
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetContainingTypeLib(IntPtr ppTLB, IntPtr pIndex)
        {
            Before($"{nameof(ppTLB)}: {ppTLB}, {nameof(pIndex)}: {pIndex}");
            var result = _wrapper.GetContainingTypeLib(ppTLB, pIndex);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetTypeAttr(IntPtr ppTypeAttr)
        {
            Before($"{nameof(ppTypeAttr)}: {ppTypeAttr}");
            var result = _wrapper.GetTypeAttr(ppTypeAttr);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetTypeComp(IntPtr ppTComp)
        {
            Before($"{nameof(ppTComp)}: {ppTComp}");
            var result = _wrapper.GetTypeComp(ppTComp);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetFuncDesc(int index, IntPtr ppFuncDesc)
        {
            Before($"{nameof(index)}: {index}, {nameof(ppFuncDesc)}: {ppFuncDesc}");
            var result = _wrapper.GetFuncDesc(index, ppFuncDesc);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetVarDesc(int index, IntPtr ppVarDesc)
        {
            Before($"{nameof(index)}: {index}, {nameof(ppVarDesc)}: {ppVarDesc}");
            var result = _wrapper.GetVarDesc(index, ppVarDesc);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetNames(int memid, IntPtr rgBstrNames, int cMaxNames, IntPtr pcNames)
        {
            Before($"{nameof(memid)}: {memid}, {nameof(rgBstrNames)}: {rgBstrNames}, {nameof(cMaxNames)}: {cMaxNames}: {nameof(pcNames)}: {pcNames}");
            var result = _wrapper.GetNames(memid, rgBstrNames, cMaxNames, pcNames);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetRefTypeOfImplType(int index, IntPtr href)
        {
            Before($"{nameof(index)}: {index}, {nameof(href)}: {href}");
            var result = _wrapper.GetRefTypeOfImplType(index, href);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetImplTypeFlags(int index, IntPtr pImplTypeFlags)
        {
            Before($"{nameof(index)}: {index}, {nameof(pImplTypeFlags)}: {pImplTypeFlags}");
            var result = _wrapper.GetImplTypeFlags(index, pImplTypeFlags);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetIDsOfNames(IntPtr rgszNames, int cNames, IntPtr pMemId)
        {
            Before($"{nameof(rgszNames)}: {rgszNames}, {nameof(cNames)}: {cNames}, {nameof(pMemId)}: {pMemId}");
            var result = _wrapper.GetIDsOfNames(rgszNames, cNames, pMemId);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int Invoke(IntPtr pvInstance, int memid, short wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo,
            IntPtr puArgErr)
        {
            Before($"parameters not supplied");
            var result = _wrapper.Invoke(pvInstance, memid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
            After($"parameters not supplied");
            return result;
        }
    }
}
