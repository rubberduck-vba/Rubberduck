using System;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// A wrapper for <see cref="ComTypes.ITypeLib"/> objects, with specific extensions for VBE hosted ITypeLibs. For usage examples, please see <see cref="VBETypeLibsAccessor"/>.
    /// </summary>
    /// <remarks>
    /// Allow safe managed consumption of VBA provided type libraries, plus exposes 
    /// a VBE extensions property for accessing VBE specific extensions.
    /// Can also be cast to <see cref="ComTypes.ITypeLib"/> for raw access to the underlying type library
    ///
    /// <see cref="ComTypes.ITypeInfo"/>s from a VBA hosted project, and obtained through
    /// <see cref="VBETypeLibsAccessor"/> will have the following behaviours:
    /// 
    ///   will expose both public and private procedures and fields
    ///   will expose constants values, but they are unnamed (their member IDs will be MEMBERID_NIL)
    ///   enumerations are not exposed directly in the type library
    ///   enumerations may be referenced by field/argument datatypes, and the ITypeInfos for them are then accessible that way
    ///   UDTs are not exposed directly in the type library
    ///   UDTs may be referenced by field/argument datatypes, and as such the ITypeInfos for them are then accessible that way
    ///   
    /// TypeInfos obtained by other means (such as the IDispatch::GetTypeInfo method) usually expose more restricted
    /// versions of ITypeInfo which may not expose private members
    /// </remarks>
    internal sealed class TypeLibWrapper : TypeLibInternalSelfMarshalForwarderBase, ITypeLibWrapper
    {
        private DisposableList<ITypeInfoWrapper> _cachedTypeInfos;
        private ComPointer<ITypeLibInternal> _typeLibPointer;

        private ITypeLibInternal _target_ITypeLib => _typeLibPointer.Interface;
        
        public bool HasVBEExtensions { get; private set; }
        public ITypeInfoWrapperCollection TypeInfos { get; private set; }

        // helpers
        public string Name => CachedTextFields._name;
        public string DocString => CachedTextFields._docString;
        public int HelpContext => CachedTextFields._helpContext;
        public string HelpFile => CachedTextFields._helpFile;
        public int TypesCount => _target_ITypeLib.GetTypeInfoCount();

        private struct TypeLibTextFields
        {
            public string _name;
            public string _docString;
            public int _helpContext;
            public string _helpFile;
        }
        private TypeLibTextFields? _cachedTextFields;
        private TypeLibTextFields CachedTextFields
        {
            get
            {
                if (_cachedTextFields.HasValue)
                {
                    return _cachedTextFields.Value;
                }

                var cache = new TypeLibTextFields();
                // as a C# caller, it's easier to work with ComTypes.ITypeLib
                ((ComTypes.ITypeLib)_target_ITypeLib).GetDocumentation((int)KnownDispatchMemberIDs.MEMBERID_NIL, out cache._name, out cache._docString, out cache._helpContext, out cache._helpFile);
                if (cache._name == null && HasVBEExtensions)
                {
                    cache._name = "[VBA.Immediate.Window]";
                }
                _cachedTextFields = cache;
                return _cachedTextFields.Value;
            }
        }

        private TypeLibVBEExtensions _vbeExtensions;
        public ITypeLibVBEExtensions VBEExtensions
        {
            get
            {
                if (_vbeExtensions != null)
                {
                    return _vbeExtensions;
                }

                if (!HasVBEExtensions)
                {
                    throw new InvalidOperationException("This TypeLib does not represent a VBE project, so does not expose VBE Extensions");
                }

                _vbeExtensions = new TypeLibVBEExtensions(this, _target_ITypeLib);
                return _vbeExtensions;
            }
        }

        private void InitCommon()
        {
            TypeInfos = new TypeInfoWrapperCollection(this);
            // ReSharper disable once SuspiciousTypeConversion.Global 
            // there is no direct implementation but it can be reached via
            // IUnknown::QueryInterface which is implicitly done as part of casting
            HasVBEExtensions = _target_ITypeLib is IVBEProject;
        }

        internal static ITypeLibWrapper FromVBProject(IVBProject vbProject)
        {
            using (var references = vbProject.References)
            {
                // Now we've got the references object, we can read the internal object structure to grab the ITypeLib
                var internalReferencesObj = StructHelper.ReadComObjectStructure<VBEReferencesObj>(references.Target);
                return TypeApiFactory.GetTypeLibWrapper(internalReferencesObj._typeLib, addRef: true);
            }
        }

        private void InitFromRawPointer(IntPtr rawObjectPtr, bool addRef)
        {
            if (!UnmanagedMemoryHelper.ValidateComObject(rawObjectPtr))
            {
                throw new ArgumentException("Expected COM object, but validation failed.");
            }

            _typeLibPointer = ComPointer<ITypeLibInternal>.GetObject(rawObjectPtr, addRef);
            InitCommon();
        }

        /// <summary>
        /// Constructor -- should be called via <see cref="TypeApiFactory"/> only.
        /// </summary>
        /// <param name="rawObjectPtr">The raw unmanaged ITypeLib pointer</param>
        /// <param name="addRef">
        /// Indicates that the pointer was obtained via unorthodox methods, such as
        /// direct memory read. Setting the parameter will effect an IUnknown::AddRef
        /// on the pointer. 
        /// </param>
        internal TypeLibWrapper(IntPtr rawObjectPtr, bool addRef)
        {
            InitFromRawPointer(rawObjectPtr, addRef);
        }
        
        private bool _isDisposed;
        public override void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            _cachedTypeInfos?.Dispose();
            _typeLibPointer.Dispose();
        }

        public int GetSafeTypeInfoByIndex(int index, out ITypeInfoWrapper outTI)
        {
            outTI = null;

            using (var typeInfoPtr = AddressableVariables.Create<IntPtr>())
            {
                var hr = _target_ITypeLib.GetTypeInfo(index, typeInfoPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr))
                {
                    return HandleBadHRESULT(hr);
                }

                var outVal = TypeApiFactory.GetTypeInfoWrapper(typeInfoPtr.Value);
                _cachedTypeInfos = _cachedTypeInfos ?? new DisposableList<ITypeInfoWrapper>();
                _cachedTypeInfos.Add(outVal);
                outTI = outVal;

                return hr;
            }
        }

        int ITypeLibWrapper.GetSafeTypeInfoByIndex(int index, out ITypeInfoWrapper outTI)
        {
            var result = GetSafeTypeInfoByIndex(index, out var outTIW);
            outTI = outTIW;
            return result;
        }

        private ComTypes.TYPELIBATTR? _cachedLibAttribs;
        public ComTypes.TYPELIBATTR Attributes
        {
            get
            {
                if (_cachedLibAttribs.HasValue)
                {
                    return _cachedLibAttribs.Value;
                }

                using (var typeLibAttributesPtr = AddressableVariables.CreatePtrTo<ComTypes.TYPELIBATTR>())
                {
                    var hr = _target_ITypeLib.GetLibAttr(typeLibAttributesPtr.Address);
                    if (ComHelper.HRESULT_FAILED(hr))
                    {
                        return _cachedLibAttribs.Value;
                    }

                    _cachedLibAttribs = typeLibAttributesPtr.Value.Value;   // dereference the ptr, then the content
                    var pTypeLibAttr = typeLibAttributesPtr.Value.Address; // dereference the ptr, and take the contents address
                    _target_ITypeLib.ReleaseTLibAttr(pTypeLibAttr);         // can release immediately as _cachedLibAttribs is a copy
                }
                return _cachedLibAttribs.Value;
            }
        }

        public IntPtr GetCOMReferencePtr()
            => RdMarshal.GetComInterfaceForObject(this, typeof(ITypeLibInternal));

        int HandleBadHRESULT(int hr)
        {
            return hr;
        }

        public override int GetTypeInfoCount()
        {
            var retVal = _target_ITypeLib.GetTypeInfoCount();
            return retVal;
        }

        public override int GetTypeInfo(int index, IntPtr ppTI)
        {
            // We have to wrap the ITypeInfo returned by GetTypeInfo
            var hr = GetSafeTypeInfoByIndex(index, out var ti);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            RdMarshal.WriteIntPtr(ppTI, ti.GetCOMReferencePtr());
            return hr;
        }
        public override int GetTypeInfoType(int index, IntPtr pTKind)
        {
            var hr = _target_ITypeLib.GetTypeInfoType(index, pTKind);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            var tKind = RdMarshal.ReadInt32(pTKind);
            tKind = (int)TypeInfoWrapper.PatchTypeKind((TYPEKIND_VBE)tKind);
            RdMarshal.WriteInt32(pTKind, tKind);

            return hr;
        }
        public override int GetTypeInfoOfGuid(ref Guid guid, IntPtr ppTInfo)
        {
            var hr = _target_ITypeLib.GetTypeInfoOfGuid(guid, ppTInfo);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            var pTInfo = RdMarshal.ReadIntPtr(ppTInfo);
            using (var outVal = TypeApiFactory.GetTypeInfoWrapper(pTInfo)) // takes ownership of the COM reference [pTInfo]
            {
                RdMarshal.WriteIntPtr(ppTInfo, outVal.GetCOMReferencePtr());

                _cachedTypeInfos = _cachedTypeInfos ?? new DisposableList<ITypeInfoWrapper>();
                _cachedTypeInfos.Add(outVal);
            }

            return hr;
        }

        public override int GetLibAttr(IntPtr ppTLibAttr)
        {
            var hr = _target_ITypeLib.GetLibAttr(ppTLibAttr);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override int GetTypeComp(IntPtr ppTComp)
        {
            var hr = _target_ITypeLib.GetTypeComp(ppTComp);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override int GetDocumentation(int index, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            var hr = _target_ITypeLib.GetDocumentation(index, strName, strDocString, dwHelpContext, strHelpFile);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override int IsName(string szNameBuf, int lHashVal, IntPtr pfName)
        {
            var hr = _target_ITypeLib.IsName(szNameBuf, lHashVal, pfName);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override int FindName(string szNameBuf, int lHashVal, IntPtr ppTInfo, IntPtr rgMemId, IntPtr pcFound)
        {
            var hr = _target_ITypeLib.FindName(szNameBuf, lHashVal, ppTInfo, rgMemId, pcFound);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override void ReleaseTLibAttr(IntPtr pTLibAttr)
        {
            _target_ITypeLib.ReleaseTLibAttr(pTLibAttr);
        }
    }
}
