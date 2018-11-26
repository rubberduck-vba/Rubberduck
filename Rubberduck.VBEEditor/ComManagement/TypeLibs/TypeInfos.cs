using System;
using System.Runtime.InteropServices;
using System.Globalization;
using Rubberduck.VBEditor.ComManagement.TypeLibsSupport;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// An enumeration used for identifying the type of a VBA document class
    /// </summary>
    public enum DocClassType
    {
        Unrecognized,
        ExcelWorkbook,
        ExcelWorksheet,
        AccessForm,
        AccessReport,
    }

    /// <summary>
    /// A class for holding known document class types used in VBA hosts, and their corresponding interface progIds
    /// </summary>
    public struct KnownDocType
    {
        public string DocTypeInterfaceProgId;
        public DocClassType DocType;

        public KnownDocType(string docTypeInterfaceProgId, DocClassType docType)
        {
            DocTypeInterfaceProgId = docTypeInterfaceProgId;
            DocType = docType;
        }
    }

    /// <summary>
    /// A helper class for providing a static array of known VBA document class types
    /// </summary>
    public static class DocClassHelper
    {
        public static KnownDocType[] KnownDocumentInterfaces =
        {
        new KnownDocType("Excel._Workbook",     DocClassType.ExcelWorkbook),
        new KnownDocType("Excel._Worksheet",    DocClassType.ExcelWorksheet),
        new KnownDocType("Access._Form",        DocClassType.AccessForm),
        new KnownDocType("Access._Form2",       DocClassType.AccessForm),
        new KnownDocType("Access._Form3",       DocClassType.AccessForm),
        new KnownDocType("Access._Report",      DocClassType.AccessReport),
        new KnownDocType("Access._Report2",     DocClassType.AccessReport),
        new KnownDocType("Access._Report3",     DocClassType.AccessReport),
    };

        // string array of the above progIDs, created once at runtime
        public static string[] KnownDocumentInterfaceProgIds;

        static DocClassHelper()
        {
            int index = 0;
            KnownDocumentInterfaceProgIds = new string[KnownDocumentInterfaces.Length];
            foreach (var knownDocClass in KnownDocumentInterfaces)
            {
                KnownDocumentInterfaceProgIds[index++] = knownDocClass.DocTypeInterfaceProgId;
            }
        }
    }

    /// <summary>
    /// A class that represents a function definition within a typeinfo
    /// </summary>
    public class TypeInfoFunc : IDisposable
    {
        private readonly TypeInfoWrapper _typeInfo;
        private IntPtr _funcDescPtr;
        private readonly string[] _names = new string[255];   // includes argument names
        private readonly int _cNames = 0;

        public ComTypes.FUNCDESC FuncDesc { get; }

        public TypeInfoFunc(TypeInfoWrapper typeInfo, int funcIndex)
        {
            _typeInfo = typeInfo;

            ((ComTypes.ITypeInfo)_typeInfo).GetFuncDesc(funcIndex, out _funcDescPtr);
            FuncDesc = StructHelper.ReadStructureUnsafe<ComTypes.FUNCDESC>(_funcDescPtr);

            ((ComTypes.ITypeInfo)_typeInfo).GetNames(FuncDesc.memid, _names, _names.Length, out _cNames);
            if (_cNames == 0) _names[0] = "[unnamed]";
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing)
            {
                return;
            }

            if (_funcDescPtr != IntPtr.Zero)
            {
                ((ComTypes.ITypeInfo)_typeInfo).ReleaseFuncDesc(_funcDescPtr);
            }
            _funcDescPtr = IntPtr.Zero;
        }

        public string Name => _names[0]; 
        public int ParamCount => FuncDesc.cParams; 

        public enum PROCKIND
        {
            PROCKIND_PROC,
            PROCKIND_LET,
            PROCKIND_SET,
            PROCKIND_GET
        }

        public PROCKIND ProcKind
        {
            get
            {
                // _funcDesc.invkind is a set of flags, and as such we convert into PROCKIND for simplicity
                if (FuncDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF))
                {
                    return PROCKIND.PROCKIND_SET;
                }
                if (FuncDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT))
                {
                    return PROCKIND.PROCKIND_LET;
                }
                if (FuncDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYGET))
                {
                    return PROCKIND.PROCKIND_GET;
                }
                return PROCKIND.PROCKIND_PROC;
            }
        }

        public void Document(StringLineBuilder output)
        {
            string namesInfo = _names[0] + "(";

            int argIndex = 1;
            while (argIndex < _cNames)
            {
                if (argIndex > 1) namesInfo += ", ";
                namesInfo += _names[argIndex].Length > 0 ? _names[argIndex] : "retVal";
                argIndex++;
            }

            namesInfo += ")";

            output.AppendLine("- member: " + namesInfo + " [id 0x" + FuncDesc.memid.ToString("X") + ", " + FuncDesc.invkind + "]");
        }
    }

    /// <summary>
    /// A class that represents a reference within a VBA project
    /// </summary>
    public class TypeInfoReference
    {
        private readonly TypeLibWrapper _vbeTypeLib;
        private readonly int _typeLibIndex;

        public string RawString { get; }
        public Guid GUID { get; }
        public uint MajorVersion { get; }
        public uint MinorVersion { get; }
        public uint LCID { get; }
        public string Path { get; }
        public string Name { get; }

        public TypeInfoReference(TypeLibWrapper vbeTypeLib, int typeLibIndex, string referenceStringRaw)
        {
            _vbeTypeLib = vbeTypeLib;
            _typeLibIndex = typeLibIndex;

            // Example: "*\G{000204EF-0000-0000-C000-000000000046}#4.1#9#C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7\VBE7.DLL#Visual Basic For Applications"
            // LibidReference defined at https://msdn.microsoft.com/en-us/library/dd922767(v=office.12).aspx
            // The string is split into 5 parts, delimited by #

            RawString = referenceStringRaw;

            var referenceStringParts = referenceStringRaw.Split(new char[] { '#' }, 5);
            if (referenceStringParts.Length != 5)
            {
                throw new ArgumentException($"Invalid reference string got {referenceStringRaw}.  Expected 5 parts.");
            }

            GUID = Guid.Parse(referenceStringParts[0].Substring(3));
            var versionSplit = referenceStringParts[1].Split(new char[] { '.' }, 2);
            if (versionSplit.Length != 2)
            {
                throw new ArgumentException($"Invalid reference string got {referenceStringRaw}.  Invalid version string.");
            }
            MajorVersion = uint.Parse(versionSplit[0], NumberStyles.AllowHexSpecifier);
            MinorVersion = uint.Parse(versionSplit[1], NumberStyles.AllowHexSpecifier);

            LCID = uint.Parse(referenceStringParts[2], NumberStyles.AllowHexSpecifier);
            Path = referenceStringParts[3];
            Name = referenceStringParts[4];
        }

        public TypeLibWrapper TypeLib => _vbeTypeLib.GetVBEReferenceTypeLibByIndex(_typeLibIndex);

        public void Document(StringLineBuilder output)
        {
            output.AppendLine("- VBE Reference: " + Name + " [path: " + Path + ", majorVersion: " + MajorVersion +
                                ", minorVersion: " + MinorVersion + ", guid: " + GUID + ", lcid: " + LCID + "]");
        }
    }

    /// <summary>
    /// A class that represents a variable or field within a typeinfo
    /// </summary>
    public class TypeInfoVar : IDisposable
    {
        private readonly TypeInfoWrapper _typeInfo;
        private readonly ComTypes.VARDESC _varDesc;
        private IntPtr _varDescPtr;
        private readonly string _name;

        public string Name => _name;

        public TypeInfoVar(TypeInfoWrapper typeInfo, int index)
        {
            _typeInfo = typeInfo;

            ((ComTypes.ITypeInfo)_typeInfo).GetVarDesc(index, out _varDescPtr);
            _varDesc = StructHelper.ReadStructureUnsafe<ComTypes.VARDESC>(_varDescPtr);

            var names = new string[1];
            if (_varDesc.memid != (int)TypeLibConsts.MEMBERID_NIL)
            {
                ((ComTypes.ITypeInfo)_typeInfo).GetNames(_varDesc.memid, names, names.Length, out _);
                _name = names[0];
            }
            else
            {
                _name = "{unknown}";     // VBA Constants appear in the typelib with no name
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            if (_varDescPtr != IntPtr.Zero)
            {
                ((ComTypes.ITypeInfo)_typeInfo).ReleaseVarDesc(_varDescPtr);
            }
            _varDescPtr = IntPtr.Zero;
            _isDisposed = true;
        }

        public void Document(StringLineBuilder output)
        {
            output.AppendLine("- field: " + _name + " [id 0x" + _varDesc.memid.ToString("X") + "]");
        }
    }

    /// <summary>
    /// A wrapper for ITypeInfo objects, with specific extensions for VBE hosted ITypeInfos
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
    /// This class can also be cast to ComTypes.ITypeInfo for raw access to the underlying type information
    /// </remarks>
    public sealed class TypeInfoWrapper : ComTypes.ITypeInfo, IDisposable
    {
        private DisposableList<TypeInfoWrapper> _typeInfosWrapped;
        private TypeLibWrapper _containerTypeLib;
        public TypeLibWrapper Container => _containerTypeLib;
        private int _containerTypeLibIndex;
        private bool _isUserFormBaseClass = false;
        private readonly IntPtr _rawObjectPtr;
        private readonly ComTypes.ITypeInfo _wrappedObjectRCW;

        private ComTypes.TYPEATTR _cachedAttributes;
        public ComTypes.TYPEATTR Attributes => _cachedAttributes;

        private readonly RestrictComInterfaceByAggregation<ComTypes.ITypeInfo> _ITypeInfo_Aggregator;
        private ComTypes.ITypeInfo target_ITypeInfo => _ITypeInfo_Aggregator?.WrappedObject ?? _wrappedObjectRCW;

        private readonly RestrictComInterfaceByAggregation<IVBEComponent> _IVBEComponent_Aggregator;
        private IVBEComponent target_IVBEComponent => _IVBEComponent_Aggregator?.WrappedObject;

        public bool HasVBEExtensions => _IVBEComponent_Aggregator?.WrappedObject != null;

        public bool HasModuleScopeCompilationErrors { get; private set; }

        /// <summary>
        /// Exposes an enumerable collection of functions provided by the ITypeInfo
        /// </summary>
        public class FuncsCollection : IIndexedCollectionBase<TypeInfoFunc>
        {
            private readonly TypeInfoWrapper _parent;
            public FuncsCollection(TypeInfoWrapper parent) => _parent = parent;
            public override int Count => _parent.Attributes.cFuncs;
            public override TypeInfoFunc GetItemByIndex(int index) => new TypeInfoFunc(_parent, index);

            public TypeInfoFunc Find(string name, TypeInfoFunc.PROCKIND procKind)
            {
                foreach (var func in this)
                {
                    if ((func.Name == name) && (func.ProcKind == procKind)) return func;
                }
                return null;
            }
        }
        public FuncsCollection Funcs;

        /// <summary>
        /// Exposes an enumerable collection of variables/fields provided by the ITypeInfo
        /// </summary>
        public class VarsCollection : IIndexedCollectionBase<TypeInfoVar>
        {
            private readonly TypeInfoWrapper _parent;
            public VarsCollection(TypeInfoWrapper parent) => _parent = parent;
            public override int Count => _parent.Attributes.cVars;
            public override TypeInfoVar GetItemByIndex(int index) => new TypeInfoVar(_parent, index);
        }
        public VarsCollection Vars;

        /// <summary>
        /// Exposes an enumerable collection of implemented interfaces provided by the ITypeInfo
        /// </summary>
        public class ImplementedInterfacesCollection : IIndexedCollectionBase<TypeInfoWrapper>
        {
            private readonly TypeInfoWrapper _parent;
            public ImplementedInterfacesCollection(TypeInfoWrapper parent) => _parent = parent;
            public override int Count => _parent.Attributes.cImplTypes;
            public override TypeInfoWrapper GetItemByIndex(int index) => _parent.GetSafeImplementedTypeInfo(index);

            /// <summary>
            /// Determines whether the type implements one of the specified interfaces
            /// </summary>
            /// <param name="interfaceProgIds">Array of interface identifiers in the format "LibName.InterfaceName"</param>
            /// <param name="matchedIndex">on return, contains the index into interfaceProgIds that matched, or -1 </param>
            /// <returns>true if the type does implement one of the specified interfaces</returns>
            public bool DoesImplement(string[] interfaceProgIds, out int matchedIndex)
            {
                matchedIndex = 0;
                foreach (var interfaceProgId in interfaceProgIds)
                {
                    if (DoesImplement(interfaceProgId))
                    {
                        return true;
                    }
                    matchedIndex++;
                }
                matchedIndex = -1;
                return false;
            }

            /// <summary>
            /// Determines whether the type implements the specified interface
            /// </summary>
            /// <param name="interfaceProgId">Interface identifier in the format "LibName.InterfaceName"</param>
            /// <returns>true if the type does implement the specified interface</returns>
            public bool DoesImplement(string interfaceProgId)
            {
                var progIdSplit = interfaceProgId.Split(new char[] { '.' }, 2);
                if (progIdSplit.Length != 2)
                {
                    throw new ArgumentException($"Expected a progid in the form of 'LibraryName.InterfaceName', got {interfaceProgId}");
                }
                return DoesImplement(progIdSplit[0], progIdSplit[1]);
            }

            /// <summary>
            /// Determines whether the type implements the specified interface
            /// </summary>
            /// <param name="containerName">The library container name</param>
            /// <param name="interfaceName">The interface name</param>
            /// <returns>true if the type does implement the specified interface</returns>
            public bool DoesImplement(string containerName, string interfaceName)
            {
                foreach (var typeInfo in this)
                {
                    using (typeInfo)
                    {
                        if ((typeInfo.Container?.Name == containerName) && (typeInfo.Name == interfaceName)) return true;
                        if (typeInfo.ImplementedInterfaces.DoesImplement(containerName, interfaceName)) return true;
                    }
                }

                return false;
            }

            /// <summary>
            /// Determines whether the type implements one of the specified interfaces
            /// </summary>
            /// <param name="interfaceIIDs">Array of interface IIDs to match</param>
            /// <param name="matchedIndex">on return, contains the index into interfaceIIDs that matched, or -1 </param>
            /// <returns>true if the type does implement one of the specified interfaces</returns>
            public bool DoesImplement(Guid[] interfaceIIDs, out int matchedIndex)
            {
                matchedIndex = 0;
                foreach (var interfaceIID in interfaceIIDs)
                {
                    if (DoesImplement(interfaceIID))
                    {
                        return true;
                    }
                    matchedIndex++;
                }
                matchedIndex = -1;
                return false;
            }

            /// <summary>
            /// Determines whether the type implements the specified interface
            /// </summary>
            /// <param name="interfaceIID">The interface IID to match</param>
            /// <returns>true if the type does implement the specified interface</returns>
            public bool DoesImplement(Guid interfaceIID)
            {
                foreach (var typeInfo in this)
                {
                    using (typeInfo)
                    {
                        if (typeInfo.GUID == interfaceIID) return true;
                        if (typeInfo.ImplementedInterfaces.DoesImplement(interfaceIID)) return true;
                    }
                }

                return false;
            }

            public TypeInfoWrapper Get(string searchTypeName)
            {
                foreach (var typeInfo in this)
                {
                    if (typeInfo.Name == searchTypeName) return typeInfo;
                    typeInfo.Dispose();
                }

                throw new ArgumentException($"TypeInfoWrapper::Get failed. '{searchTypeName}' component not found.");
            }
        }
        public ImplementedInterfacesCollection ImplementedInterfaces;

        private void InitCommon()
        {
            Funcs = new FuncsCollection(this);
            Vars = new VarsCollection(this);
            ImplementedInterfaces = new ImplementedInterfacesCollection(this);

            IntPtr typeAttrPtr = IntPtr.Zero;
            try
            {
                target_ITypeInfo.GetTypeAttr(out typeAttrPtr);
                _cachedAttributes = StructHelper.ReadStructureUnsafe<ComTypes.TYPEATTR>(typeAttrPtr);
                target_ITypeInfo.ReleaseTypeAttr(typeAttrPtr);      // don't need to keep a hold of it, as _cachedAttributes is a copy
            }
            catch (Exception e)
            {
                // If there is a compilation error outside of a procedure code block, the type information is not available for that component.
                // We detect this, via the E_VBA_COMPILEERROR error 
                if (e.HResult == (int)KnownComHResults.E_VBA_COMPILEERROR)
                {
                    HasModuleScopeCompilationErrors = true;
                }

                // just mute the erorr and expose an empty type
                _cachedAttributes = new ComTypes.TYPEATTR();
            }

            // cache the container type library if it is available
            try
            {
                // We have to wrap the ITypeLib returned by GetContainingTypeLib
                // so we cast to our ITypeInfo_Ptrs interface in order to work with the raw IntPtrs
                ((ITypeInfo_Ptrs)target_ITypeInfo).GetContainingTypeLib(out var typeLibPtr, out _containerTypeLibIndex);
                _containerTypeLib?.Dispose();
                _containerTypeLib = new TypeLibWrapper(typeLibPtr, true);  
            }
            catch (Exception e)
            {
                if (e.HResult == (int)KnownComHResults.E_NOTIMPL)
                {
                    // it is acceptable for a type to not have a container, as types can be runtime generated (e.g. UserForm base classes)
                    // When that is the case, the ITypeInfo responds with E_NOTIMPL
                }
                else
                {
                    throw new ArgumentException("Unrecognised error when getting ITypeInfo container: \n" + e.ToString());
                }
            }
        }

        public TypeInfoWrapper(ComTypes.ITypeInfo rawTypeInfo)
        {
            _wrappedObjectRCW = rawTypeInfo;
            InitCommon();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rawObjectPtr">The raw unmanaged pointer to the ITypeInfo</param>
        /// <param name="parentUserFormUniqueId">used internally for providing a name for UserForm base classes</param>
        public TypeInfoWrapper(IntPtr rawObjectPtr, int? parentUserFormUniqueId = null)
        {
            _rawObjectPtr = rawObjectPtr;
            if (_rawObjectPtr == IntPtr.Zero)
            {
                throw new ArgumentException("Unepectedly received a null pointer.");
            }

            // We have to restrict interface requests to VBE hosted ITypeInfos due to a bug in their implementation.
            // See TypeInfoWrapper class XML doc for details.

            // queryForType is passed as false for ITypeInfo here, as rawObjectPtr is known to point to the ITypeInfo vtable
            // additionally allowing it to query for ITypeInfo gives a _different_, and more prohibitive implementation of ITypeInfo
            _ITypeInfo_Aggregator = new RestrictComInterfaceByAggregation<ComTypes.ITypeInfo>(rawObjectPtr, queryForType: false);
            _IVBEComponent_Aggregator = new RestrictComInterfaceByAggregation<IVBEComponent>(rawObjectPtr);
            
            // base classes of VBE UserForms cause an access violation on calling GetDocumentation(MEMBERID_NIL)
            // so we have to detect UserForm parents, and ensure GetDocumentation(MEMBERID_NIL) never gets through
            if (parentUserFormUniqueId.HasValue)
            {
                _cachedTextFields = new TypeLibTextFields { _name = "_UserFormBase{unnamed}#" + parentUserFormUniqueId };
            }

            InitCommon();
            DetectUserFormClass();
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            _typeInfosWrapped?.Dispose();
            _containerTypeLib?.Dispose();
            _ITypeInfo_Aggregator?.Dispose();
            _IVBEComponent_Aggregator?.Dispose();
            
            Marshal.Release(_rawObjectPtr);
        }

        private TypeLibTextFields? _cachedTextFields;

        private TypeLibTextFields CachedTextFields
        {
            get
            {
                if (!_cachedTextFields.HasValue)
                {
                    var cache = new TypeLibTextFields();
                    target_ITypeInfo.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out cache._name, out cache._docString, out cache._helpContext, out cache._helpFile);
                    _cachedTextFields = cache;
                }
                return _cachedTextFields.Value;
            }
        }

        public string Name => CachedTextFields._name;
        public string DocString => CachedTextFields._docString;
        public int HelpContext => CachedTextFields._helpContext;
        public string HelpFile => CachedTextFields._helpFile;

        public string GetProgID() => (Container?.Name ?? "{unnamedlibrary}") + "." + CachedTextFields._name;

        public Guid GUID => Attributes.guid;
        public TYPEKIND_VBE TypeKind => (TYPEKIND_VBE)Attributes.typekind;

        public bool HasPredeclaredId => Attributes.wTypeFlags.HasFlag(ComTypes.TYPEFLAGS.TYPEFLAG_FPREDECLID);
        public ComTypes.TYPEFLAGS Flags => Attributes.wTypeFlags;

        private bool HasNoContainer() => _containerTypeLib == null;

        /// <summary>
        /// Silently compiles the individual VBE component (class/module etc)
        /// </summary>
        /// <returns>true if this module, plus any direct dependent modules compile successfully</returns>
        public bool CompileComponent()
        {
            if (!HasVBEExtensions)
            {
                throw new ArgumentException("This TypeInfo does not represent a VBE component, so we cannot compile it");
            }

            try
            {
                target_IVBEComponent.CompileComponent();
                return true;
            }
            catch (Exception e)
            {
#if DEBUG
                if (e.HResult != (int)KnownComHResults.E_VBA_COMPILEERROR)
                {
                    // When debugging we want to know if there are any other errors returned by the compiler as
                    // the error code might be useful.
                    throw new ArgumentException("Unrecognised VBE compiler error: \n" + e.ToString());
                }
#endif

                return false;
            }
        }

        /// <summary>
        /// Used to detect UserForm classes, needed to workaround a VBE bug.  See TypeInfoWrapper XML doc for details. 
        /// </summary>
        private void DetectUserFormClass()
        {
            // Determine if this is a UserForm base class, that requires special handling to workaround a VBE bug in its implemented classes
            // the guids are dynamic, so we can't use them for detection.
            if ((TypeKind == TYPEKIND_VBE.TKIND_COCLASS) &&
                    HasNoContainer() &&
                    (ImplementedInterfaces.Count == 2) &&
                    (Name == "Form"))
            {
                // we can be 99.999999% sure it IS the runtime generated UserForm base class
                _isUserFormBaseClass = true;
            }
        }

        /// <summary>
        /// Provides an accessor object for invoking methods on a standard module in a VBA project
        /// </summary>
        /// <remarks>caller is responsible for calling ReleaseComObject</remarks>
        /// <returns>the accessor object</returns>
        public IDispatch GetStdModAccessor()
        {
            if (!HasVBEExtensions)
            {
                throw new ArgumentException("This ITypeInfo is not hosted by the VBE, so does not support GetStdModAccessor");
            }

            return target_IVBEComponent.GetStdModAccessor();
        }

        /// <summary>
        /// Executes a procedure inside a standard module in a VBA project
        /// </summary>
        /// <param name="name">the name of the procedure to invoke</param>
        /// <param name="args">arguments to pass to the procedure</param>
        /// <remarks>the returned object can be a COM object, and the callee is responsible for releasing it appropriately</remarks>
        /// <returns>an object representing the return value from the procedure, or null if none.</returns>
        public object StdModExecute(string name, object[] args = null)
        {
            if (!HasVBEExtensions)
            {
                throw new ArgumentException("This ITypeInfo is not hosted by the VBE, so does not support StdModExecute");
            }

            // We search for the dispId using the real type info rather than using staticModule.GetIdsOfNames, 
            // as we can then also include PRIVATE scoped procedures.
            var func = Funcs.Find(name, TypeInfoFunc.PROCKIND.PROCKIND_PROC);
            if (func == null)
            {
                throw new ArgumentException($"StdModExecute failed.  Couldn't find procedure named '{name}'");
            }

            var staticModule = GetStdModAccessor();

            try
            {
                return IDispatchHelper.Invoke(staticModule, func.FuncDesc.memid, IDispatchHelper.InvokeKind.DISPATCH_METHOD, args);
            }
            finally
            {
                Marshal.ReleaseComObject(staticModule);
            }
        }

        /// <summary>
        /// Gets the control ITypeInfo by looking for the corresponding getter on the form interface and returning its retval type
        /// </summary>
        /// <param name="controlName">the name of the control</param>
        /// <returns>TypeInfoWrapper representing the type of control, typically the coclass, but this is host dependent</returns>
        public TypeInfoWrapper GetControlType(string controlName)
        {
            // TODO should encapsulate handling of raw datatypes
            foreach (var func in Funcs)
            {
                using (func)
                {
                    // Controls are exposed as getters on the interface.
                    //     can either be    ControlType* get_ControlName()       
                    //     or               HRESULT get_ControlName(ControlType** Out) 

                    if ((func.Name == controlName) &&
                        (func.ProcKind == TypeInfoFunc.PROCKIND.PROCKIND_GET) &&
                        (func.ParamCount == 0) &&
                        (func.FuncDesc.elemdescFunc.tdesc.vt == (short)VarEnum.VT_PTR))
                    {
                        var retValElement = StructHelper.ReadStructureUnsafe<ComTypes.ELEMDESC>(func.FuncDesc.elemdescFunc.tdesc.lpValue);
                        if (retValElement.tdesc.vt == (short)VarEnum.VT_USERDEFINED)
                        {
                            return GetSafeRefTypeInfo((int)retValElement.tdesc.lpValue);
                        }
                    }
                    else if ((func.Name == controlName) &&
                        (func.ProcKind == TypeInfoFunc.PROCKIND.PROCKIND_GET) &&
                        (func.ParamCount == 1) &&
                        (func.FuncDesc.elemdescFunc.tdesc.vt == (short)VarEnum.VT_HRESULT))
                    {
                        // Get details of the first argument
                        var retValElementOuterPtr = StructHelper.ReadStructureUnsafe<ComTypes.ELEMDESC>(func.FuncDesc.lprgelemdescParam);
                        if (retValElementOuterPtr.tdesc.vt == (short)VarEnum.VT_PTR)
                        {
                            var retValElementInnerPtr = StructHelper.ReadStructureUnsafe<ComTypes.ELEMDESC>(retValElementOuterPtr.tdesc.lpValue);
                            if (retValElementInnerPtr.tdesc.vt == (short)VarEnum.VT_PTR)
                            {
                                var retValElement = StructHelper.ReadStructureUnsafe<ComTypes.ELEMDESC>(retValElementInnerPtr.tdesc.lpValue);

                                if (retValElement.tdesc.vt == (short)VarEnum.VT_USERDEFINED)
                                {
                                    return GetSafeRefTypeInfo((int)retValElement.tdesc.lpValue);
                                }
                            }
                        }
                    }
                }
            }

            throw new ArgumentException($"TypeInfoWrapper::GetControlType failed. '{controlName}' control not found.");
        }

        public TypeInfoWrapper GetSafeRefTypeInfo(int hRef)
        {
            // we cast to our ITypeInfo_Ptrs interface in order to work with the raw IntPtr for aggregation
            ((ITypeInfo_Ptrs)target_ITypeInfo).GetRefTypeInfo(hRef, out var typeInfoPtr);
            var outVal = new TypeInfoWrapper(typeInfoPtr, _isUserFormBaseClass ? (int?)hRef : null); // takes ownership of the COM reference

            _typeInfosWrapped?.Dispose();
            _typeInfosWrapped = _typeInfosWrapped ?? new DisposableList<TypeInfoWrapper>();
            _typeInfosWrapped.Add(outVal);
            return outVal;
        }

        public TypeInfoWrapper GetSafeImplementedTypeInfo(int index)
        {
            target_ITypeInfo.GetRefTypeOfImplType(index, out var href);
            return GetSafeRefTypeInfo(href);
        }

        /// <summary>
        /// Determines the document class type of a VBA class.  See DocClassHelper
        /// </summary>
        /// <returns>the identified document class type, or DocClassType.Unrecognized</returns>
        public DocClassType DetermineDocumentClassType()
        {
            if (ImplementedInterfaces.DoesImplement(DocClassHelper.KnownDocumentInterfaceProgIds, out int matchId))
            {
                return DocClassHelper.KnownDocumentInterfaces[matchId].DocType;
            }
            return DocClassType.Unrecognized;
        }

        public void Document(StringLineBuilder output, string qualifiedName, int implementsLevel)
        {
            output.AppendLine();
            if (implementsLevel == 0)
            {
                output.AppendLine("-------------------------------------------------------------------------------");
                output.AppendLine();
            }
            implementsLevel++;

            qualifiedName += "::" + (Name ?? "[unnamed]");
            output.AppendLineNoNullChars(qualifiedName);
            output.AppendLineNoNullChars("- Documentation: " + DocString);
            output.AppendLineNoNullChars("- HelpContext: " + HelpContext);
            output.AppendLineNoNullChars("- HelpFile: " + HelpFile);

            output.AppendLine("- HasVBEExtensions: " + HasVBEExtensions);
            if (HasVBEExtensions) output.AppendLine("- HasModuleScopeCompilationErrors: " + HasModuleScopeCompilationErrors);

            output.AppendLine("- Type: " + TypeKind);
            output.AppendLine("- Guid: {" + GUID + "}");

            output.AppendLine("- cImplTypes (implemented interfaces count): " + ImplementedInterfaces.Count);
            output.AppendLine("- cFuncs (function count): " + Funcs.Count);
            output.AppendLine("- cVars (fields count): " + Vars.Count);

            foreach (var func in Funcs)
            {
                using (func)
                {
                    func.Document(output);
                }
            }
            foreach (var variable in Vars)
            {
                using (variable)
                {
                    variable.Document(output);
                }
            }
            foreach (var typeInfoImpl in ImplementedInterfaces)
            {
                using (typeInfoImpl)
                {
                    output.AppendLine("implements...");
                    typeInfoImpl.Document(output, qualifiedName, implementsLevel);
                }
            }
        }

        // And finally we act as a safe pass-through to the raw ITypeInfo interface
        // We have to wrap all ITypeInfos
        void ComTypes.ITypeInfo.GetRefTypeInfo(int hRef, out ComTypes.ITypeInfo ppTI)
            => ppTI = GetSafeRefTypeInfo(hRef);
        void ComTypes.ITypeInfo.GetContainingTypeLib(out ComTypes.ITypeLib ppTLB, out int pIndex)
        {
            ppTLB = _containerTypeLib;
            pIndex = _containerTypeLibIndex;
        }
        void ComTypes.ITypeInfo.GetTypeAttr(out IntPtr ppTypeAttr)
            => target_ITypeInfo.GetTypeAttr(out ppTypeAttr);
        void ComTypes.ITypeInfo.GetTypeComp(out ComTypes.ITypeComp ppTComp)
            => target_ITypeInfo.GetTypeComp(out ppTComp);
        void ComTypes.ITypeInfo.GetFuncDesc(int index, out IntPtr ppFuncDesc)
            => target_ITypeInfo.GetFuncDesc(index, out ppFuncDesc);
        void ComTypes.ITypeInfo.GetVarDesc(int index, out IntPtr ppVarDesc)
            => target_ITypeInfo.GetVarDesc(index, out ppVarDesc);
        void ComTypes.ITypeInfo.GetNames(int memid, string[] rgBstrNames, int cMaxNames, out int pcNames)
            => target_ITypeInfo.GetNames(memid, rgBstrNames, cMaxNames, out pcNames);
        void ComTypes.ITypeInfo.GetRefTypeOfImplType(int index, out int href)
            => target_ITypeInfo.GetRefTypeOfImplType(index, out href);
        void ComTypes.ITypeInfo.GetImplTypeFlags(int index, out ComTypes.IMPLTYPEFLAGS pImplTypeFlags)
            => target_ITypeInfo.GetImplTypeFlags(index, out pImplTypeFlags);
        void ComTypes.ITypeInfo.GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId)
            => target_ITypeInfo.GetIDsOfNames(rgszNames, cNames, pMemId);
        void ComTypes.ITypeInfo.Invoke(object pvInstance, int memid, short wFlags, ref ComTypes.DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr)
            => target_ITypeInfo.Invoke(pvInstance, memid, wFlags, ref pDispParams, pVarResult, pExcepInfo, out puArgErr);
        void ComTypes.ITypeInfo.GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile)
        {
            if (index == (int)TypeLibConsts.MEMBERID_NIL)
            {
                // return the cached information here, to workaround the VBE bug for unnamed UserForm base classes causing an access violation
                strName = Name;
                strDocString = DocString;
                dwHelpContext = HelpContext;
                strHelpFile = HelpFile;
            }
            else
            {
                target_ITypeInfo.GetDocumentation(index, out strName, out strDocString, out dwHelpContext, out strHelpFile);
            }
        }
        void ComTypes.ITypeInfo.GetDllEntry(int memid, ComTypes.INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
            => target_ITypeInfo.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
        void ComTypes.ITypeInfo.AddressOfMember(int memid, ComTypes.INVOKEKIND invKind, out IntPtr ppv)
            => target_ITypeInfo.AddressOfMember(memid, invKind, out ppv);
        void ComTypes.ITypeInfo.CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj)
            => target_ITypeInfo.CreateInstance(pUnkOuter, riid, out ppvObj);
        void ComTypes.ITypeInfo.GetMops(int memid, out string pBstrMops)
            => target_ITypeInfo.GetMops(memid, out pBstrMops);
        void ComTypes.ITypeInfo.ReleaseTypeAttr(IntPtr pTypeAttr)
            => target_ITypeInfo.ReleaseTypeAttr(pTypeAttr);
        void ComTypes.ITypeInfo.ReleaseFuncDesc(IntPtr pFuncDesc)
            => target_ITypeInfo.ReleaseFuncDesc(pFuncDesc);
        void ComTypes.ITypeInfo.ReleaseVarDesc(IntPtr pVarDesc)
            => target_ITypeInfo.ReleaseVarDesc(pVarDesc);
    }
}