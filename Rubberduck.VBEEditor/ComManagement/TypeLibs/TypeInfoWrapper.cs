using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// An extended version of TYPEKIND, which VBA uses internally to identify VBA classes as a seperate type
    /// see https://msdn.microsoft.com/en-us/library/windows/desktop/ms221643(v=vs.85).aspx
    /// </summary>
    public enum TYPEKIND_VBE
    {
        TKIND_ENUM = 0,
        TKIND_RECORD = 1,
        TKIND_MODULE = 2,
        TKIND_INTERFACE = 3,
        TKIND_DISPATCH = 4,
        TKIND_COCLASS = 5,
        TKIND_ALIAS = 6,
        TKIND_UNION = 7,

        TKIND_VBACLASS = 8,
    }

    /// <summary>
    /// A wrapper for <see cref="ComTypes.ITypeInfo"/> objects, with specific
    /// extensions for VBE hosted ITypeInfos
    /// </summary>
    /// <remarks>
    /// There are two significant bugs in the VBE implementations for ITypeInfo that we have to work around.
    /// 1)  Some implementations of ITypeInfo provided by the VBE will crash with an AV if you call 
    ///      IProvideClassInfo::GetClassInfo on them.  And guess what method the CLR calls on all COM interop objects 
    ///      when creating a RCW?  You guessed it.  So, we use an aggregation object, plus ITypeInfo and ITypeLib wrappers 
    ///      to circumvent this VBE bug.
    ///
    /// 2)  The ITypeInfo for base classes of UserForms crash with an AV if you call ITypeInfo::GetDocumentation(MEMBERID_NIL) 
    ///     to get the type name.  We've got to remember that the VBE didn't ever intend for us to get hold of these objects, 
    ///     so there will be little bugs.  This bug is also resolved in the provided wrappers.
    ///
    /// This class can also be cast to <see cref="ComTypes.ITypeInfo"/> for safe access to the underlying type information
    ///
    /// TypeInfos from a VBA hosted project, and obtained through <see cref="VBETypeLibsAccessor"/> will have the following behaviours:
    /// 
    ///   will expose both public and private procedures and fields
    ///   will expose constants values, but they are unnamed (their member IDs will be MEMBERID_NIL)
    ///   module level variables in a standard module are hidden (ref to <see cref="TypeInfoConstantsCollection"/> for details)
    ///   enumerations are not exposed directly in the type library
    ///   enumerations may be referenced by field/argument datatypes, and the ITypeInfos for them are then accessible that way
    ///   UDTs are not exposed directly in the type library
    ///   UDTs may be referenced by field/argument datatypes, and as such the ITypeInfos for them are then accessible that way
    ///   
    /// TypeInfos obtained by other means (such as the IDispatch::GetTypeInfo method) usually expose more restricted
    /// versions of ITypeInfo which may not expose private members
    /// </remarks>
    internal sealed class TypeInfoWrapper : TypeInfoInternalSelfMarshalForwarderBase, ITypeInfoWrapper
    {
        private DisposableList<ITypeInfoWrapper> _cachedReferencedTypeInfos;
        private ComPointer<ITypeInfoInternal> _typeInfoPointer;
        private ComPointer<ITypeInfoInternal> _typeInfoAlternatePointer;

        private ITypeInfoInternal _target_ITypeInfo => _typeInfoPointer.Interface;
        private ITypeInfoInternal _target_ITypeInfoAlternate => _typeInfoAlternatePointer.Interface;

        private ComTypes.ITypeLib _container { get; set; }
        public ComTypes.ITypeLib Container => (ComTypes.ITypeLib)_container;

        public int ContainerIndex { get; private set; }
        public bool HasModuleScopeCompilationErrors { get; private set; }
        public bool HasVBEExtensions { get; private set; }
        public ComTypes.TYPEATTR CachedAttributes { get; private set; }
        public bool HasSimulatedContainer { get; private set; }
        public bool IsUserFormBaseClass { get; private set; }

        public ITypeInfoFunctionCollection Funcs { get; private set; }

        /// <summary>
        /// A VBA module's <see cref="ComTypes.ITypeInfo"/> violates the type API conventions
        /// because a module is not supposed to have fields/variables enumerated. To
        /// work around this, we will maintain this separate lists containing only
        /// constants. For any kind of ITypeInfos, this will/should not be used.
        /// Rather use <see cref="Consts"/> which should implement the class
        /// <see cref="TypeInfoConstantsCollection"/> to help with this problem.
        /// </summary>
        public ITypeInfoVariablesCollection AllVars { get; private set; }

        /// <summary>
        /// Enumerates constants only. This is populated only for non-compliant type infos
        /// (e.g. VBA standard modules).
        /// </summary>
        /// <remarks>
        /// See <see cref="AllVars"/> for extra considerations with VBA's ITypeInfo
        /// </remarks>
        public ITypeInfoVariablesCollection Consts => _consts;
        private TypeInfoConstantsCollection _consts;

        public ITypeInfoVariablesCollection Vars
        {
            get
            {
                if (_consts != null && _consts.Count > 0)
                {
                    return _consts;
                }

                return AllVars;
            }
        }

        public ITypeInfoImplementedInterfacesCollection ImplementedInterfaces { get; private set; }

        // some helpers
        public string Name => CachedTextFields._name;
        public string DocString => CachedTextFields._docString;
        public int HelpContext => CachedTextFields._helpContext;
        public string HelpFile => CachedTextFields._helpFile;
        public string ProgID => ContainerName + "." + CachedTextFields._name;
        public Guid GUID => CachedAttributes.guid;
        public TYPEKIND_VBE TypeKind => (TYPEKIND_VBE)CachedAttributes.typekind;
        public bool HasPredeclaredId => CachedAttributes.wTypeFlags.HasFlag(ComTypes.TYPEFLAGS.TYPEFLAG_FPREDECLID);
        public ComTypes.TYPEFLAGS Flags => CachedAttributes.wTypeFlags;
        public string ContainerName => RdMarshal.GetTypeLibName(_container);
        
        // Constants inside VBA components are exposed via the ITypeInfo, but there names are not reported correctly.
        // Their names all appear with a DispID of MEMBERID_NIL.  In order to try to make VBA type infos more agreeable to the specifications,
        // we make up some unique names for these constants, and create unique DispIDs for them at runtime.  
        // This will help some ITypeInfo consumers that may not like the unnamed fields. 
        // Currently this is achieved by defining a range in the 32-bit space used by DispIDs that is unlikely to conflict with 
        // any normal DispIDs assigned by VBA, and one would think unlikely to be used by custom VB_UserMemId attributes. 
        // The range chosen allows for 65536 constants, starting at _ourConstantsDispatchMemberIDRangeStart.
        // generated names are in the format "_constantFieldId" + Index (where index is the index into GetVarDesc)
        const int _ourConstantsDispatchMemberIDRangeStart = unchecked((int)0xFEDC0000);
        const int _ourConstantsDispatchMemberIDRangeBitmaskCheck = unchecked((int)0xFFFF0000);
        const int _ourConstantsDispatchMemberIDIndexBitmask = unchecked((int)0x0000FFFF);

        private static bool IsDispatchMemberIDInOurConstantsRange(int memid)
        {
            return (memid & _ourConstantsDispatchMemberIDRangeBitmaskCheck) == _ourConstantsDispatchMemberIDRangeStart;
        }

        private void InitCommon()
        {
            using (var typeAttrPtr = AddressableVariables.CreatePtrTo<ComTypes.TYPEATTR>())
            {
                var hr = _target_ITypeInfo.GetTypeAttr(typeAttrPtr.Address);

                if (!ComHelper.HRESULT_FAILED(hr))
                {
                    CachedAttributes = typeAttrPtr.Value.Value;     // dereference the ptr, then the content
                    var pTypeAttr = typeAttrPtr.Value.Address;     // dereference the ptr, and take the contents address
                    _target_ITypeInfo.ReleaseTypeAttr(pTypeAttr);   // can release immediately as CachedAttributes is a copy
                }
                else
                {
                    if (hr == (int)KnownComHResults.E_VBA_COMPILEERROR)
                    {
                        // If there is a compilation error outside of a procedure code block, the type information is not available for that component.
                        // We detect this, via the E_VBA_COMPILEERROR error 
                        HasModuleScopeCompilationErrors = true;
                    }

                    // just mute the error and expose an empty type
                    CachedAttributes = new ComTypes.TYPEATTR();
                }
            }

            Funcs = new TypeInfoFunctionCollection(this, CachedAttributes);

            // Refer to AllVars XML docs for details
            AllVars = new TypeInfoVariablesCollection(this, CachedAttributes);
            if (CachedAttributes.typekind == ComTypes.TYPEKIND.TKIND_MODULE && HasVBEExtensions)
            {
                _consts = new TypeInfoConstantsCollection(this, CachedAttributes);
            }

            ImplementedInterfaces = new TypeInfoImplementedInterfacesCollection(this, CachedAttributes);

            // cache the container type library if it is available, else create a simulated one
            using (var typeLibPtr = AddressableVariables.Create<IntPtr>())
            using (var containerTypeLibIndex = AddressableVariables.Create<int>())
            {
                var hr = _target_ITypeInfo.GetContainingTypeLib(typeLibPtr.Address, containerTypeLibIndex.Address);

                if (!ComHelper.HRESULT_FAILED(hr))
                {
                    // We have to wrap the ITypeLib returned by GetContainingTypeLib
                    _container = TypeApiFactory.GetTypeLibWrapper(typeLibPtr.Value, addRef: false);
                    ContainerIndex = containerTypeLibIndex.Value;
                }
                else
                {
                    if (hr == (int)KnownComHResults.E_NOTIMPL)
                    {
                        // it is acceptable for a type to not have a container, as types can be runtime generated (e.g. UserForm base classes)
                        // When that is the case, the ITypeInfo responds with E_NOTIMPL
                        // However, we create fake container to avoid errors from CLR when using those "uncontained" TypeInfo
                        HasSimulatedContainer = true;
                        var newContainer = new SimpleCustomTypeLibrary();
                        _container = newContainer;
                        ContainerIndex = newContainer.Add(this);
                    }
                    else
                    {
                        throw new ArgumentException("Unrecognised error when getting ITypeInfo container: \n" + hr);
                    }
                }
            }
        }

        private void InitFromRawPointer(IntPtr rawObjectPtr, bool addRef)
        {
            if (!UnmanagedMemoryHelper.ValidateComObject(rawObjectPtr))
            {
                throw new ArgumentException("Expected COM object, but validation failed.");
            }

            // We have to restrict interface requests to VBE hosted ITypeInfos due to a bug in their implementation.
            // See TypeInfoWrapper class XML doc for details.

            // VBE provides two implementations of ITypeInfo for each component.  Both versions have different quirks and limitations.
            // We use both versions to try to expose a more complete/accurate version of ITypeInfo.
            _typeInfoPointer =
                ComPointer<ITypeInfoInternal>.GetObjectViaAggregation(rawObjectPtr, addRef, queryType: false);
            _typeInfoAlternatePointer =
                ComPointer<ITypeInfoInternal>.GetObjectViaAggregation(rawObjectPtr, addRef, queryType: true);
            
            // safely test whether the provided ITypeInfo is hosted by the VBE, and thus exposes the VBE extensions
            HasVBEExtensions = ComHelper.DoesComObjPtrSupportInterface<IVBEComponent>(rawObjectPtr);
            
            InitCommon();
            DetectUserFormClass();
        }

        /// <summary>
        /// Constructor -- should be called via <see cref="TypeApiFactory"/> only.
        /// </summary>
        /// <param name="rawTypeInfo">The ITypeInfo object to be wrapped. It will be checked if it's already wrapped to avoid double-wraping.</param>
        internal TypeInfoWrapper(ComTypes.ITypeInfo rawTypeInfo)
        {
            if ((rawTypeInfo as TypeInfoInternalSelfMarshalForwarderBase) != null)
            {
                // The passed in TypeInfo is already a TypeInfoWrapper.  Detect & prevent double wrapping...
                var tlib = (TypeInfoWrapper)(TypeInfoInternalSelfMarshalForwarderBase)rawTypeInfo;
                var rawObjectPtr = tlib._typeInfoPointer.ExtractPointer();
                InitFromRawPointer(rawObjectPtr, addRef: true);
                _cachedTextFields = tlib._cachedTextFields;     // copied to ensure we work around the UserForm GetDocumentation() crash
                return;
            }

            InitCommon();
        }

        /// <summary>
        /// Constructor -- should be called via <see cref="TypeApiFactory"/> only.
        /// </summary>
        /// <param name="rawObjectPtr">The raw unmanaged pointer to the ITypeInfo.  This class takes ownership, and will call Marshall.Release() on it upon disposal.</param>
        /// <param name="parentUserFormUniqueId">used internally for providing a name for UserForm base classes</param>
        internal TypeInfoWrapper(IntPtr rawObjectPtr, int? parentUserFormUniqueId = null)
        {
            // base classes of VBE UserForms cause an access violation on calling GetDocumentation(MEMBERID_NIL)
            // so we have to detect UserForm parents, and ensure GetDocumentation(MEMBERID_NIL) never gets through
            // we do that by caching the GetDocumentation(MEMBERID_NIL) result into _cachedTextFields, or overriding it here
            if (parentUserFormUniqueId.HasValue)
            {
                _cachedTextFields = new TypeLibTextFields { _name = "_UserFormBase{unnamed}#" + parentUserFormUniqueId };
            }

            InitFromRawPointer(rawObjectPtr, addRef: false);
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            _vbeExtensions?.Dispose();
            _cachedReferencedTypeInfos?.Dispose();
            (_container as IDisposable)?.Dispose();

            _typeInfoPointer.Dispose();
            _typeInfoAlternatePointer?.Dispose();
        }

        TypeInfoVBEExtensions _vbeExtensions;
        public ITypeInfoVBEExtensions VBEExtensions
        {
            get
            {
                if (_vbeExtensions != null)
                {
                    return _vbeExtensions;
                }

                if (!HasVBEExtensions) throw new InvalidOperationException("This TypeInfo does not represent a VBE component, so does not expose VBE Extensions");
                _vbeExtensions = new TypeInfoVBEExtensions(this, _typeInfoPointer.ExtractPointer());

                return _vbeExtensions;
            }
        }

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
                ((ComTypes.ITypeInfo)_target_ITypeInfo).GetDocumentation((int)KnownDispatchMemberIDs.MEMBERID_NIL, out cache._name, out cache._docString, out cache._helpContext, out cache._helpFile);
                _cachedTextFields = cache;
                return _cachedTextFields.Value;
            }
        }

        public static ComTypes.TYPEKIND PatchTypeKind(TYPEKIND_VBE typeKind)
        {
            // We patch up the special TKIND_VBACLASS constant to TKIND_DISPATCH as that seems the most appropriate
            // supporting both variables[fields] and functions[members]
            if (typeKind == TYPEKIND_VBE.TKIND_VBACLASS)
            {
                return ComTypes.TYPEKIND.TKIND_DISPATCH;
            }
            return (ComTypes.TYPEKIND)typeKind;
        }

        /// <summary>
        /// Used to detect UserForm classes, needed to workaround a VBE bug.  See <cref see="TypeInfoWrapper"> for details. 
        /// </summary>
        private void DetectUserFormClass()
        {
            // Determine if this is a UserForm base class, that requires special handling to workaround a VBE bug in its implemented classes
            // the guids are dynamic, so we can't use them for detection.
            if ((TypeKind == TYPEKIND_VBE.TKIND_COCLASS) &&
                    HasSimulatedContainer &&
                    (ImplementedInterfaces.Count == 2) &&
                    (Name == "Form"))
            {
                // we can be 99.999999% sure it IS the runtime generated UserForm base class
                IsUserFormBaseClass = true;
            }
        }

        public int GetSafeRefTypeInfo(int hRef, out ITypeInfoWrapper outTI)
        {
            outTI = null;

            using (var typeInfoPtr = AddressableVariables.Create<IntPtr>())
            {
                var hr = _target_ITypeInfo.GetRefTypeInfo(hRef, typeInfoPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

                var outVal = TypeApiFactory.GetTypeInfoWrapper(typeInfoPtr.Value, IsUserFormBaseClass ? (int?)hRef : null); // takes ownership of the COM reference
                _cachedReferencedTypeInfos = _cachedReferencedTypeInfos ?? new DisposableList<ITypeInfoWrapper>();
                _cachedReferencedTypeInfos.Add(outVal);
                outTI = outVal;

                return hr;
            }
        }

        int ITypeInfoWrapper.GetSafeRefTypeInfo(int hRef, out ITypeInfoWrapper outTI)
        {
            var result = GetSafeRefTypeInfo(hRef, out var outTIW);
            outTI = outTIW;
            return result;
        }

        public IntPtr GetCOMReferencePtr()
            => RdMarshal.GetComInterfaceForObject(this, typeof(ITypeInfoInternal));

        private static int HandleBadHRESULT(int hr)
        {
            return hr;
        }

        public override int GetContainingTypeLib(IntPtr ppTLB, IntPtr pIndex)
        {
            // even though pIndex is described as a non-optional OUT argument, mscorlib sometimes calls this with a nullptr from the C++ side.
            if (pIndex == IntPtr.Zero)
            {
                RdMarshal.WriteIntPtr(ppTLB, IntPtr.Zero);
                return (int)KnownComHResults.E_INVALIDARG;
            }

            RdMarshal.WriteIntPtr(ppTLB, RdMarshal.GetComInterfaceForObject(_container, typeof(ITypeLibInternal)));
            if (pIndex != IntPtr.Zero) RdMarshal.WriteInt32(pIndex, ContainerIndex);

            return (int)KnownComHResults.S_OK;
        }

        public override int GetTypeAttr(IntPtr ppTypeAttr)
        {
            var hr = _target_ITypeInfo.GetTypeAttr(ppTypeAttr);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            var pTypeAttr = StructHelper.ReadStructureUnsafe<IntPtr>(ppTypeAttr);
            var typeAttr = StructHelper.ReadStructureUnsafe<ComTypes.TYPEATTR>(pTypeAttr);

            typeAttr.typekind = PatchTypeKind((TYPEKIND_VBE)typeAttr.typekind);

            // A proper type library would not have any variables defined in the modules. 
            // However, a VBA module can have those and that can trip up CLR's parser. To
            // work around this, we lie to any ComTypes.ITypeInfo consumer that there's no 
            // variables. If we need for VBA-specific work, we need to consult
            // ITypeInfoWrapper.Vars instead.
            if (typeAttr.typekind == ComTypes.TYPEKIND.TKIND_MODULE && typeAttr.cVars > 0)
            {
                typeAttr.cVars = (short)Consts.Count;
            }

            RdMarshal.StructureToPtr<ComTypes.TYPEATTR>(typeAttr, pTypeAttr, true);
            return hr;
        }

        public override int GetTypeComp(IntPtr ppTComp)
        {
            var hr = _target_ITypeInfo.GetTypeComp(ppTComp);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override int GetFuncDesc(int index, IntPtr ppFuncDesc)
        {
            var hr = _target_ITypeInfo.GetFuncDesc(index, ppFuncDesc);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            if (_target_ITypeInfoAlternate != null)
            {
                var pFuncDesc = StructHelper.ReadStructureUnsafe<IntPtr>(ppFuncDesc);
                var funcDesc = StructHelper.ReadStructureUnsafe<ComTypes.FUNCDESC>(pFuncDesc);

                // Populate wFuncFlags from the alternative typeinfo provided by VBA
                // The alternative typeinfo is not as useful as the main typeinfo for most things, but does expose wFuncFlags
                // The list of functions appears to be in the same order as the main typeinfo.  
                using (var funcDescAlternatePtr = AddressableVariables.CreatePtrTo<ComTypes.FUNCDESC>())
                {
                    var hr2 = _target_ITypeInfoAlternate.GetFuncDesc(index, funcDescAlternatePtr.Address);
                    if (!ComHelper.HRESULT_FAILED(hr2))
                    {
                        var funcDescAlternate = funcDescAlternatePtr.Value.Value;    // dereference the ptr, then the content

                        //sanity check
                        if (funcDescAlternate.memid == funcDesc.memid)
                        {
                            funcDesc.wFuncFlags = funcDescAlternate.wFuncFlags;
                        }
                        else
                        {
                            Debug.Assert(false, $"The sanity check failed; {nameof(funcDesc.memid)}: {funcDesc.memid} and {nameof(funcDescAlternate.memid)}: {funcDescAlternate.memid}");
                        }
                        _target_ITypeInfoAlternate.ReleaseFuncDesc(funcDescAlternatePtr.Value.Address);

                        RdMarshal.StructureToPtr(funcDesc, pFuncDesc, true);
                    }
                }
            }

            return hr;
        }

        public override int GetVarDesc(int index, IntPtr ppVarDesc)
        {
            var safeIndex = _consts != null && _consts.Count > 0 
                ? _consts.MappedIndex(index) 
                : index;

            var hr = _target_ITypeInfo.GetVarDesc(safeIndex, ppVarDesc);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            var pVarDesc = StructHelper.ReadStructureUnsafe<IntPtr>(ppVarDesc);
            var varDesc = StructHelper.ReadStructureUnsafe<ComTypes.VARDESC>(pVarDesc);
            if (varDesc.memid == (int)KnownDispatchMemberIDs.MEMBERID_NIL)
            {
                // constants are not reported correctly in VBA type infos.  They all have MEMBERID_NIL set.
                // we will provide fake DispIds and names to satisfy parsers.  Shit but works for now.
                varDesc.memid = (int)(_ourConstantsDispatchMemberIDRangeStart + safeIndex);
                RdMarshal.StructureToPtr(varDesc, pVarDesc, true);
            }
            else
            {
                // Unlike GetFuncDesc, we can't get the wVarFlags for fields from the alternative VBA ITypeInfo
                // because GetVarDesc() hard crashes on the alternative typeinfo
                /*
                    if (target_ITypeInfoAlternate != null)
                    {
                        using (var varDescPtr2 = AddressableVariables.CreatePtrTo<ComTypes.VARDESC>())
                        {
                            var hr2 = target_ITypeInfoAlternate.GetVarDesc(index, varDescPtr2.Address);
                            var varDesc2 = varDescPtr2.Value.Value; // dereference the ptr, then the content
                            VarDesc.wVarFlags = varDesc2.wVarFlags;
                            RdMarshal.StructureToPtr(VarDesc, pVarDesc, false);
                        }
                   }
                */
            }
            return hr;
        }

        public override int GetNames(int memid, IntPtr rgBstrNames, int cMaxNames, IntPtr pcNames)
        {
            if (IsDispatchMemberIDInOurConstantsRange(memid))
            {
                // this is most likely one of our simulated names from GetVarDesc()
                var fieldId = memid & _ourConstantsDispatchMemberIDIndexBitmask;
                if ((rgBstrNames != IntPtr.Zero) && (cMaxNames >= 1))
                {
                    // output 1 string to the array
                    RdMarshal.WriteIntPtr(rgBstrNames, RdMarshal.StringToBSTR("_constantFieldId" + fieldId));
                    if (pcNames != IntPtr.Zero) RdMarshal.WriteInt32(pcNames, 1);
                    return (int)KnownComHResults.S_OK;
                }
            }

            var hr = _target_ITypeInfo.GetNames(memid, rgBstrNames, cMaxNames, pcNames);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override int GetRefTypeOfImplType(int index, IntPtr href)
        {
            var hr = _target_ITypeInfo.GetRefTypeOfImplType(index, href);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override int GetImplTypeFlags(int index, IntPtr pImplTypeFlags)
        {
            var hr = _target_ITypeInfo.GetImplTypeFlags(index, pImplTypeFlags);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override int GetIDsOfNames(IntPtr rgszNames, int cNames, IntPtr pMemId)
        {
            var hr = _target_ITypeInfo.GetIDsOfNames(rgszNames, cNames, pMemId);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override int Invoke(IntPtr pvInstance, int memid, short wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr)
        {
            var hr = _target_ITypeInfo.Invoke(pvInstance, memid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override int GetDocumentation(int memid, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            if (IsDispatchMemberIDInOurConstantsRange(memid))
            {
                // this is very likely one of our simulated names from GetVarDesc()
                var fieldId = memid & _ourConstantsDispatchMemberIDIndexBitmask;
                if (strName != IntPtr.Zero) RdMarshal.WriteIntPtr(strName, Marshal.StringToBSTR("_constantFieldId" + fieldId));
                if (strDocString != IntPtr.Zero) RdMarshal.WriteIntPtr(strDocString, IntPtr.Zero);
                if (dwHelpContext != IntPtr.Zero) RdMarshal.WriteInt32(dwHelpContext, 0);
                if (strHelpFile != IntPtr.Zero) RdMarshal.WriteIntPtr(strHelpFile, IntPtr.Zero);
                return (int)KnownComHResults.S_OK;
            }

            if (memid == (int)KnownDispatchMemberIDs.MEMBERID_NIL)
            {
                // return the cached information here, to workaround the VBE bug for unnamed UserForm base classes causing an access violation
                if (strName != IntPtr.Zero) RdMarshal.WriteIntPtr(strName, RdMarshal.StringToBSTR(Name));
                if (strDocString != IntPtr.Zero) RdMarshal.WriteIntPtr(strDocString, RdMarshal.StringToBSTR(DocString));
                if (dwHelpContext != IntPtr.Zero) RdMarshal.WriteInt32(dwHelpContext, HelpContext);
                if (strHelpFile != IntPtr.Zero) RdMarshal.WriteIntPtr(strHelpFile, RdMarshal.StringToBSTR(HelpFile));
                return (int)KnownComHResults.S_OK;
            }
            else
            {
                var hr = _target_ITypeInfo.GetDocumentation(memid, strName, strDocString, dwHelpContext, strHelpFile);
                return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
            }
        }

        public override int GetDllEntry(int memid, ComTypes.INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
        {
            var hr = _target_ITypeInfo.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override int GetRefTypeInfo(int hRef, IntPtr ppTI)
        {
            var hr = GetSafeRefTypeInfo(hRef, out var ti);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            RdMarshal.WriteIntPtr(ppTI, ti.GetCOMReferencePtr());
            return hr;
        }

        public override int AddressOfMember(int memid, ComTypes.INVOKEKIND invKind, IntPtr ppv)
        {
            var hr = _target_ITypeInfo.AddressOfMember(memid, invKind, ppv);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override int CreateInstance(IntPtr pUnkOuter, ref Guid riid, IntPtr ppvObj)
        {
            var hr = _target_ITypeInfo.CreateInstance(pUnkOuter, riid, ppvObj);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override int GetMops(int memid, IntPtr pBstrMops)
        {
            var hr = _target_ITypeInfo.GetMops(memid, pBstrMops);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }

        public override void ReleaseTypeAttr(IntPtr pTypeAttr)
        {
            _target_ITypeInfo.ReleaseTypeAttr(pTypeAttr);
        }

        public override void ReleaseFuncDesc(IntPtr pFuncDesc)
        {
            _target_ITypeInfo.ReleaseFuncDesc(pFuncDesc);
        }

        public override void ReleaseVarDesc(IntPtr pVarDesc)
        {
            _target_ITypeInfo.ReleaseVarDesc(pVarDesc);
        }
    }
}