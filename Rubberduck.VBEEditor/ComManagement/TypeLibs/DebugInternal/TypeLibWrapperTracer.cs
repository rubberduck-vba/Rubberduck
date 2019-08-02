using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.DebugInternal
{
    internal class TypeLibWrapperTracer : ITypeLibWrapper, ITypeLibInternal
    {
        private readonly ITypeLibWrapper _wrapper;
        private readonly ITypeLibInternal _inner;

        internal TypeLibWrapperTracer(ITypeLibWrapper wrapper, ITypeLibInternal inner)
        {
            _wrapper = wrapper;
            _inner = inner;
        }

        private static void Before(string parameters = null, [CallerMemberName] string methodName = null)
        {
            System.Diagnostics.Debug.Print($"Entering {nameof(ITypeLibWrapper)}::{methodName}; {parameters}");
        }

        private static void After(string parameters = null, [CallerMemberName] string methodName = null)
        {
            System.Diagnostics.Debug.Print($"Leaving {nameof(ITypeLibWrapper)}::{methodName}; {parameters}");
        }

        public int GetTypeInfoCount()
        {
            Before();
            var result = _wrapper.GetTypeInfoCount();
            After($"{nameof(result)}: {result}");
            return result;
        }

        public int GetTypeInfo(int index, IntPtr ppTI)
        {
            Before($"{nameof(index)}: {index}, {nameof(ppTI)}: {ppTI}");
            var result = _inner.GetTypeInfo(index, ppTI);
            After($"{nameof(result)}: {result}, {nameof(ppTI)}: {ppTI}");
            return result;
        }

        public int GetTypeInfoType(int index, IntPtr pTKind)
        {
            Before($"{nameof(index)}: {index}, {nameof(pTKind)}: {pTKind}");
            var result = _inner.GetTypeInfoType(index, pTKind);
            After($"{nameof(result)}: {result}, {nameof(pTKind)}: {pTKind}");
            return result;
        }

        public int GetTypeInfoOfGuid(ref Guid guid, IntPtr ppTInfo)
        {
            Before($"{nameof(guid)}: {guid}, {nameof(ppTInfo)}: {ppTInfo}");
            var result = _inner.GetTypeInfoOfGuid(ref guid, ppTInfo);
            After($"{nameof(result)}: {result}, {nameof(ppTInfo)}: {ppTInfo}");
            return result;
        }

        public int GetLibAttr(IntPtr ppTLibAttr)
        {
            Before($"{nameof(ppTLibAttr)}: {ppTLibAttr}");
            var result = _inner.GetLibAttr(ppTLibAttr);
            After($"{nameof(result)}: {result}, {nameof(ppTLibAttr)}: {ppTLibAttr}");
            return result;
        }

        public int GetTypeComp(IntPtr ppTComp)
        {
            Before($"{nameof(ppTComp)}: {ppTComp}");
            var result = _inner.GetTypeComp(ppTComp);
            After($"{nameof(result)}: {result}, {nameof(ppTComp)}: {ppTComp}");
            return result;
        }

        public int GetDocumentation(int index, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            Before($"{nameof(index)}: {index}, {nameof(strName)}: {strName}, {nameof(strDocString)}: {strDocString}, {nameof(dwHelpContext)}: {dwHelpContext}, {nameof(strHelpFile)}: {strHelpFile}");
            var result = _inner.GetDocumentation(index, strName, strDocString, dwHelpContext, strHelpFile);
            After($"{nameof(result)}: {result}, {nameof(strName)}: {strName}, {nameof(strDocString)}: {strDocString}, {nameof(dwHelpContext)}: {dwHelpContext}, {nameof(strHelpFile)}: {strHelpFile}");
            return result;
        }

        public int IsName(string szNameBuf, int lHashVal, IntPtr pfName)
        {
            Before($"{nameof(szNameBuf)}: {szNameBuf}, {nameof(lHashVal)}: {lHashVal}, {nameof(pfName)}: {pfName}");
            var result = _inner.IsName(szNameBuf, lHashVal, pfName);
            After($"{nameof(result)}: {result}, {nameof(pfName)}: {pfName}");
            return result;
        }

        public int FindName(string szNameBuf, int lHashVal, IntPtr ppTInfo, IntPtr rgMemId, IntPtr pcFound)
        {
            Before($"{nameof(szNameBuf)}: {szNameBuf}, {nameof(lHashVal)}: {lHashVal}, {nameof(ppTInfo)}: {ppTInfo}, {nameof(rgMemId)}: {rgMemId}, {nameof(pcFound)}: {pcFound}");
            var result = _inner.FindName(szNameBuf, lHashVal, ppTInfo, rgMemId, pcFound);
            After($"{nameof(result)}: {result}, {nameof(ppTInfo)}: {ppTInfo}, {nameof(rgMemId)}: {rgMemId}, {nameof(pcFound)}: {pcFound}");
            return result;
        }

        public void GetTypeInfo(int index, out ITypeInfo ppTI)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetTypeInfo(index, out var t);
            After($"{nameof(ppTI)}: {t?.GetHashCode()}");
            ppTI = t;
        }

        public void GetTypeInfoType(int index, out TYPEKIND pTKind)
        {
            Before($"{nameof(index)}: {index}");
            _wrapper.GetTypeInfoType(index, out var t);
            After($"{nameof(pTKind)}: {t}");
            pTKind = t;
        }

        public void GetTypeInfoOfGuid(ref Guid guid, out ITypeInfo ppTInfo)
        {
            Before($"{nameof(guid)}: {guid}");
            _wrapper.GetTypeInfoOfGuid(ref guid, out var t);
            After($"{nameof(ppTInfo)}: {t?.GetHashCode()}");
            ppTInfo = t;
        }

        public void GetLibAttr(out IntPtr ppTLibAttr)
        {
            Before();
            _wrapper.GetLibAttr(out var t);
            After($"{nameof(ppTLibAttr)}: {t}");
            ppTLibAttr = t;
        }

        public void GetTypeComp(out ITypeComp ppTComp)
        {
            Before();
            _wrapper.GetTypeComp(out var t);
            After($"{nameof(ppTComp)}: {t?.GetHashCode()}");
            ppTComp = t;
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

        public bool IsName(string szNameBuf, int lHashVal)
        {
            Before($"{nameof(szNameBuf)}: {szNameBuf}, {nameof(lHashVal)}: {lHashVal}");
            var result = _wrapper.IsName(szNameBuf, lHashVal);
            After($"{nameof(result)}: {result}");
            return result;
        }

        public void FindName(string szNameBuf, int lHashVal, ITypeInfo[] ppTInfo, int[] rgMemId, ref short pcFound)
        {
            Before($"{nameof(szNameBuf)}: {szNameBuf}, {nameof(lHashVal)}: {lHashVal}, {ppTInfo}: {(ppTInfo == null ? "null" : "objects")}, {nameof(rgMemId)}: {(rgMemId == null ? "null" : "ints")}, {nameof(pcFound)}: {pcFound}");
            _wrapper.FindName(szNameBuf, lHashVal, ppTInfo, rgMemId, ref pcFound);
            After($"{ppTInfo}: {(ppTInfo == null ? "null" : "objects")}, {nameof(rgMemId)}: {(rgMemId == null ? "null" : "ints")}, {nameof(pcFound)}: {pcFound}");
        }

        public void ReleaseTLibAttr(IntPtr pTLibAttr)
        {
            Before($"{nameof(pTLibAttr)}: {pTLibAttr}");
            _wrapper.ReleaseTLibAttr(pTLibAttr);
            After();
        }

        public void Dispose()
        {
            Before();
            _wrapper.Dispose();
            After();
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

        public int TypesCount
        {
            get
            {
                Before();
                var result = _wrapper.TypesCount;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public ITypeInfoWrapperCollection TypeInfos
        {
            get
            {
                Before();
                var result = _wrapper.TypeInfos;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public ITypeLibVBEExtensions VBEExtensions
        {
            get
            {
                Before();
                var result = _wrapper.VBEExtensions;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public TYPELIBATTR Attributes
        {
            get
            {
                Before();
                var result = _wrapper.Attributes;
                After($"{nameof(result)}: {result}");
                return result;
            }
        }

        public int GetSafeTypeInfoByIndex(int index, out ITypeInfoWrapper outTI)
        {
            Before($"{nameof(index)}: {index}");
            var result = _wrapper.GetSafeTypeInfoByIndex(index, out var t);
            After($"{nameof(result)}: {result}, {nameof(outTI)}: {t?.GetHashCode()}");
            outTI = t;
            return result;
        }
    }
}
