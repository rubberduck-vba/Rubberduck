using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Utility
{
    /// <summary>
    /// A custom implementation of <see cref="ComTypes.ITypeLib"/>.
    /// A dumb container of ITypeInfos.
    /// </summary>
    /// <remarks>
    /// The TypeInfos it holds are disposed of elsewhere.
    /// (typically in an earlier typeinfo references collection) 
    /// </remarks>
    internal sealed class SimpleCustomTypeLibrary : TypeLibInternalSelfMarshalForwarderBase
    {
        private readonly List<TypeInfoWrapper> _containedTypeInfos = new List<TypeInfoWrapper>();
        ComTypes.TYPELIBATTR _libAttribs;

        public SimpleCustomTypeLibrary()
        {
            // create a standard TYPELIBATTR structure, with a random runtime generated GUID that we'll return in GetLibAttr()
            _libAttribs = new ComTypes.TYPELIBATTR();
            _libAttribs.guid = Guid.NewGuid();
            _libAttribs.lcid = CultureInfo.CurrentCulture.LCID;
            _libAttribs.syskind = (IntPtr.Size == 8) ? ComTypes.SYSKIND.SYS_WIN64 : ComTypes.SYSKIND.SYS_WIN32;
            _libAttribs.wLibFlags = ComTypes.LIBFLAGS.LIBFLAG_FRESTRICTED | ComTypes.LIBFLAGS.LIBFLAG_FHIDDEN;
            _libAttribs.wMajorVerNum = 1;
        }

        private bool _isDisposed;
        public override void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
        }

        // returns the index of the new entry
        public int Add(TypeInfoWrapper ti)
        {
            _containedTypeInfos.Add(ti);
            return _containedTypeInfos.Count - 1;
        }

        public override int GetTypeInfoCount()
            => _containedTypeInfos.Count;

        public override int GetTypeInfo(int index, IntPtr ppTI)
        {
            if (index >= _containedTypeInfos.Count) return (int)KnownComHResults.TYPE_E_ELEMENTNOTFOUND;

            var ti = _containedTypeInfos[index];
            RdMarshal.WriteIntPtr(ppTI, ti.GetCOMReferencePtr());
            return (int)KnownComHResults.S_OK;
        }

        public override int GetTypeInfoType(int index, IntPtr pTKind)
        {
            if (index >= _containedTypeInfos.Count) return (int)KnownComHResults.TYPE_E_ELEMENTNOTFOUND;

            var ti = _containedTypeInfos[index];

            var typeKind = TypeInfoWrapper.PatchTypeKind(ti.TypeKind);
            RdMarshal.WriteInt32(pTKind, (int)typeKind);

            return (int)KnownComHResults.S_OK;
        }

        public override int GetTypeInfoOfGuid(ref Guid guid, IntPtr ppTInfo)
        {
            var inGuid = guid;
            var ti = _containedTypeInfos.Find(x => x.GUID == inGuid);
            if (ti == null) return (int)KnownComHResults.TYPE_E_ELEMENTNOTFOUND;

            RdMarshal.WriteIntPtr(ppTInfo, ti.GetCOMReferencePtr());
            return (int)KnownComHResults.S_OK;
        }

        public override int GetLibAttr(IntPtr ppTLibAttr)
        {
            var output = RdMarshal.AllocHGlobal(Marshal.SizeOf(typeof(ComTypes.TYPELIBATTR)));
            RdMarshal.StructureToPtr(_libAttribs, output, false);
            RdMarshal.WriteIntPtr(ppTLibAttr, output);
            return (int)KnownComHResults.S_OK;
        }

        // not important to implement for us
        public override int GetTypeComp(IntPtr ppTComp)
            => (int)KnownComHResults.E_NOTIMPL;

        public override int GetDocumentation(int memid, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            if (memid == (int)KnownDispatchMemberIDs.MEMBERID_NIL)
            {
                if (strName != IntPtr.Zero) RdMarshal.WriteIntPtr(strName, Marshal.StringToBSTR("_ArtificialContainer"));
                if (strDocString != IntPtr.Zero) RdMarshal.WriteIntPtr(strDocString, IntPtr.Zero);
                if (dwHelpContext != IntPtr.Zero) RdMarshal.WriteInt32(dwHelpContext, 0);
                if (strHelpFile != IntPtr.Zero) RdMarshal.WriteIntPtr(strHelpFile, IntPtr.Zero);

                return (int)KnownComHResults.S_OK;
            }
            return (int)KnownComHResults.TYPE_E_ELEMENTNOTFOUND;
        }

        // not important to implement for us
        public override int IsName(string szNameBuf, int lHashVal, IntPtr pfName)
            => (int)KnownComHResults.E_NOTIMPL;

        // not important to implement for us
        public override int FindName(string szNameBuf, int lHashVal, IntPtr ppTInfo, IntPtr rgMemId, IntPtr pcFound)
            => (int)KnownComHResults.E_NOTIMPL;

        public override void ReleaseTLibAttr(IntPtr pTLibAttr)
        {
            if (pTLibAttr != IntPtr.Zero) RdMarshal.FreeHGlobal(pTLibAttr);
        }
    }
}
