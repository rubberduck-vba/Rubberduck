using System;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using ComTypes = System.Runtime.InteropServices.ComTypes;

/// <summary>
/// For usage examples, please see VBETypeLibsAPI
/// </summary>
/// <remarks>
/// TypeInfos from a VBA hosted project, and obtained through VBETypeLibsAccessor will have the following behaviours:
/// 
///   will expose both public and private prcoedures and fields
///   will expose constants values, but they are unnamed (their member IDs will be MEMBERID_NIL)
///   enumerations are not exposed directly in the type library
///   enumerations may be referenced by field/argument datatypes, and the ITypeInfos for them are then accessible that way
///   UDTs are not exposed directly in the type library
///   UDTs may be referenced by field/argument datatypes, and as such the ITypeInfos for them are then accessible that way
///   
/// TypeInfos obtained by other means (such as the IDispatch::GetTypeInfo method) usually expose more restricted
/// versions of ITypeInfo which may not expose private members
/// </remarks>

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// A wrapper for ITypeLib objects, with specific extensions for VBE hosted ITypeLibs
    /// </summary>
    /// <remarks>
    /// allow safe managed consumption, plus adds ConditionalCompilationArguments property, 
    /// VBEReferences collection, and CompileProject method.
    /// Can also be cast to ComTypes.ITypeLib for raw access to the underlying type library
    /// </remarks>
    public sealed class TypeLibWrapper : ITypeLibInternalSelfMarshalForwarder, ITypeLibWrapper
    {
        private DisposableList<TypeInfoWrapper> _cachedTypeInfos;
        private IntPtr _target_ITypeLibPtr;
        private ITypeLibInternal _target_ITypeLib;
        private bool _target_ITypeLib_IsRefCounted;

        public bool HasVBEExtensions { get; private set; }
        public TypeInfoWrapperCollection TypeInfos { get; private set; }

        // helpers
        public string Name => CachedTextFields._name;
        public string DocString => CachedTextFields._docString;
        public int HelpContext => CachedTextFields._helpContext;
        public string HelpFile => CachedTextFields._helpFile;
        public int TypesCount => _target_ITypeLib.GetTypeInfoCount();

        public struct TypeLibTextFields
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
                if (!_cachedTextFields.HasValue)
                {
                    var cache = new TypeLibTextFields();
                    // as a C# caller, it's easier to work with ComTypes.ITypeLib
                    ((ComTypes.ITypeLib)_target_ITypeLib).GetDocumentation((int)KnownDispatchMemberIDs.MEMBERID_NIL, out cache._name, out cache._docString, out cache._helpContext, out cache._helpFile);
                    _cachedTextFields = cache;
                }
                return _cachedTextFields.Value;
            }
        }

        TypeLibVBEExtensions _vbeExtensions;
        public TypeLibVBEExtensions VBEExtensions
        {
            get
            {
                if (_vbeExtensions == null)
                {
                    if (!HasVBEExtensions) throw new InvalidOperationException("This TypeLib does not represent a VBE project, so does not expose VBE Extensions");
                    _vbeExtensions = new TypeLibVBEExtensions(this, _target_ITypeLib);
                }

                return _vbeExtensions;
            }
        }

        private void InitCommon()
        {
            TypeInfos = new TypeInfoWrapperCollection(this);
            HasVBEExtensions = (_target_ITypeLib as IVBEProject) != null;
        }

        internal static TypeLibWrapper FromVBProject(IVBProject vbProject)
        {
            using (var references = vbProject.References)
            {
                // Now we've got the references object, we can read the internal object structure to grab the ITypeLib
                var internalReferencesObj = StructHelper.ReadComObjectStructure<VBEReferencesObj>(references.Target);
                return new TypeLibWrapper(internalReferencesObj._typeLib, makeCopyOfReference: false);
            }
        }

        private void InitFromRawPointer(IntPtr rawObjectPtr, bool makeCopyOfReference)
        {
            if (!UnmanagedMemHelper.ValidateComObject(rawObjectPtr))
            {
                throw new ArgumentException("Expected COM object, but validation failed.");
            }
            if (makeCopyOfReference) Marshal.AddRef(rawObjectPtr);

            _target_ITypeLibPtr = rawObjectPtr;
            _target_ITypeLib = (ITypeLibInternal)Marshal.GetObjectForIUnknown(rawObjectPtr);
            _target_ITypeLib_IsRefCounted = true;
            InitCommon();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rawObjectPtr">The raw unamanaged ITypeLib pointer</param>
        public TypeLibWrapper(IntPtr rawObjectPtr, bool makeCopyOfReference)
        {
            InitFromRawPointer(rawObjectPtr, makeCopyOfReference);
        }

        public TypeLibWrapper(ComTypes.ITypeLib unwrappedTypeLib)
        {
            if ((unwrappedTypeLib as ITypeLibInternalSelfMarshalForwarder) != null)
            {
                // The passed in TypeInfo is already a TypeInfoWrapper.  Detect & prevent double wrapping...
                var tlib = (TypeLibWrapper)(ITypeLibInternalSelfMarshalForwarder)unwrappedTypeLib;
                InitFromRawPointer(tlib._target_ITypeLibPtr, makeCopyOfReference: true);
                return;
            }

            _target_ITypeLibPtr = Marshal.GetIUnknownForObject(unwrappedTypeLib);
            _target_ITypeLib = (ITypeLibInternal)unwrappedTypeLib;
            _target_ITypeLib_IsRefCounted = false;
            InitCommon();
        }

        private bool _isDisposed;
        public override void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            _cachedTypeInfos?.Dispose();
            if (_target_ITypeLib_IsRefCounted) Marshal.ReleaseComObject(_target_ITypeLib);
            if (_target_ITypeLibPtr != IntPtr.Zero) Marshal.Release(_target_ITypeLibPtr);
        }

        public int GetSafeTypeInfoByIndex(int index, out TypeInfoWrapper outTI)
        {
            outTI = null;

            using (var typeInfoPtr = AddressableVariables.Create<IntPtr>())
            {
                int hr = _target_ITypeLib.GetTypeInfo(index, typeInfoPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

                var outVal = new TypeInfoWrapper(typeInfoPtr.Value);
                _cachedTypeInfos = _cachedTypeInfos ?? new DisposableList<TypeInfoWrapper>();
                _cachedTypeInfos.Add(outVal);
                outTI = outVal;

                return hr;
            }
        }

        private ComTypes.TYPELIBATTR? _cachedLibAttribs;
        public ComTypes.TYPELIBATTR Attributes
        {
            get
            {
                if (!_cachedLibAttribs.HasValue)
                {
                    using (var typeLibAttributesPtr = AddressableVariables.CreatePtrTo<ComTypes.TYPELIBATTR>())
                    {
                        int hr = _target_ITypeLib.GetLibAttr(typeLibAttributesPtr.Address);
                        if (!ComHelper.HRESULT_FAILED(hr))
                        {
                            _cachedLibAttribs = typeLibAttributesPtr.Value.Value;   // dereference the ptr, then the content
                            var pTypeLibAttr = typeLibAttributesPtr.Value.Address; // dereference the ptr, and take the contents address
                            _target_ITypeLib.ReleaseTLibAttr(pTypeLibAttr);         // can release immediately as _cachedLibAttribs is a copy
                        }
                    }
                }
                return _cachedLibAttribs.Value;
            }
        }

        public IntPtr GetCOMReferencePtr()
            => Marshal.GetComInterfaceForObject(this, typeof(ITypeLibInternal));

        int HandleBadHRESULT(int hr)
        {
            return hr;
        }

        public override int GetTypeInfoCount()
        {
            int retVal = _target_ITypeLib.GetTypeInfoCount();
            return retVal;
        }

        public override int GetTypeInfo(int index, IntPtr ppTI)
        {
            // We have to wrap the ITypeInfo returned by GetTypeInfo
            int hr = GetSafeTypeInfoByIndex(index, out var ti);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            Marshal.WriteIntPtr(ppTI, ti.GetCOMReferencePtr());
            return hr;
        }
        public override int GetTypeInfoType(int index, IntPtr pTKind)
        {
            int hr = _target_ITypeLib.GetTypeInfoType(index, pTKind);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            var tKind = Marshal.ReadInt32(pTKind);
            tKind = (int)TypeInfoWrapper.PatchTypeKind((TYPEKIND_VBE)tKind);
            Marshal.WriteInt32(pTKind, tKind);

            return hr;
        }
        public override int GetTypeInfoOfGuid(ref Guid guid, IntPtr ppTInfo)
        {
            int hr = _target_ITypeLib.GetTypeInfoOfGuid(guid, ppTInfo);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            var pTInfo = Marshal.ReadIntPtr(ppTInfo);
            using (var outVal = new TypeInfoWrapper(pTInfo)) // takes ownership of the COM reference [pTInfo]
            {
                Marshal.WriteIntPtr(ppTInfo, outVal.GetCOMReferencePtr());

                _cachedTypeInfos = _cachedTypeInfos ?? new DisposableList<TypeInfoWrapper>();
                _cachedTypeInfos.Add(outVal);
            }

            return hr;
        }

        public override int GetLibAttr(IntPtr ppTLibAttr)
        {
            int hr = _target_ITypeLib.GetLibAttr(ppTLibAttr);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);
            return hr;
        }
        public override int GetTypeComp(IntPtr ppTComp)
        {
            int hr = _target_ITypeLib.GetTypeComp(ppTComp);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);
            return hr;
        }
        public override int GetDocumentation(int index, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            int hr = _target_ITypeLib.GetDocumentation(index, strName, strDocString, dwHelpContext, strHelpFile);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);
            return hr;
        }
        public override int IsName(string szNameBuf, int lHashVal, IntPtr pfName)
        {
            int hr = _target_ITypeLib.IsName(szNameBuf, lHashVal, pfName);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);
            return hr;
        }
        public override int FindName(string szNameBuf, int lHashVal, IntPtr ppTInfo, IntPtr rgMemId, IntPtr pcFound)
        {
            int hr = _target_ITypeLib.FindName(szNameBuf, lHashVal, ppTInfo, rgMemId, pcFound);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);
            return hr;
        }
        public override void ReleaseTLibAttr(IntPtr pTLibAttr)
        {
            _target_ITypeLib.ReleaseTLibAttr(pTLibAttr);
        }
    }
}
