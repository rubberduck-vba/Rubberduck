using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibsAbstract;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Reflection = System.Reflection;


// TODO comments/XML doc
// TODO a few FIXMEs

// make GetControlType support Access forms etc
// IsAccessForm example
// split into TypeInfos.cs

/*VBETypeLibsAPI::GetModuleFlags(ide, projectName, moduleName) to get the TYPEFLAGS
VBETypeLibsAPI::GetMemberId(ide, projectName, moduleName, memberName)
VBETypeLibsAPI::GetMemberHelpString(ide, projectName, moduleName, memberName)
*/


// USAGE GUIDE:   see class VBETypeLibsAPI for demonstrations of usage.
//
// The root object for exposure of the type libraries is VBETypeLibsAccessor.  It takes the VBE in its construtor.
// The main wrappers (TypeLibWrapper and TypeInfoWrapper) can be used as regular ITypeLib and ITypeInfo objects through casting.
//
// THIS IS A WORK IN PROGRESS.  ERROR HANDLING NEEDS WORK, AS DOES A FEW OF THE HELPER ROUTINES.
//
// WARNING: when using VBETypeLibsAccessor directly, do not cache it  
//   The type library is LIVE information, so consider it a snapshot at that very moment when you are dealing with it
//   Make sure you call VBETypeLibsAccessor.Dispose() as soon as you have done what you need to do with it.
//   Once control returns back to the VBE, you must assume that all the ITypeLib/ITypeInfo pointers are now invalid.
//
// CURRENT LIMITATIONS:
// At the moment, enums and UDTs are not exposed through the type libraries
// Constants names are not available

// IMPLEMENTATION DETAILS:
// There are two significant bugs in the VBE typeinfos implementations that we have to work around.
// 1)  Some implementations of ITypeInfo provided by the VBE will crash with an AV if you call IProvideClassInfo::GetClassInfo on them.
//      And guess what method the CLR calls on all COM interop objects when creating a RCW?  You guessed it.
//      So, we use an aggregation object, plus ITypeInfo and ITypeLib wrappers to circumvent this VBE bug.
//
// 2)  The ITypeInfo for base classes of UserForms crash with an AV if you call ITypeInfo::GetDocumentation(MEMBERID_NIL) to get the type name
//     We've got to remember that the VBE didn't ever intend for us to get hold of these objects, so there will be little bugs.
//     This bug is also resolved in the provided wrappers.
//
// All the extended functionality is exposed through the wrappers.

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    public class StringLineBuilder
    {
        StringBuilder _document = new StringBuilder();

        public override string ToString() => _document.ToString();

        public void AppendLine(string value = "")
            => _document.Append(value + "\r\n");

        public void AppendLineNoNullChars(string value)
            => AppendLine(value.Replace("\0", string.Empty));
    }

    public static class StructHelper
    {
        public static T ReadComObjectStructure<T>(object comObj)
        {
            // Reads a COM object as a structure to copy its internal fields
            if (Marshal.IsComObject(comObj))
            {
                var referencesPtr = Marshal.GetIUnknownForObject(comObj);
                var retVal = StructHelper.ReadStructure<T>(referencesPtr);
                Marshal.Release(referencesPtr);
                return retVal;
            }
            else
            {
                throw new ArgumentException("Expected a COM object");
            }
        }

        public static T ReadStructure<T>(IntPtr memAddress)
        {
            if (memAddress == IntPtr.Zero) return default(T);
            return (T)Marshal.PtrToStructure(memAddress, typeof(T));
        }

        public static T ReadStructureSafe<T>(IntPtr memAddress)
        {
            if (memAddress == IntPtr.Zero) return default(T);

            // FIXME add memory address validation here, using VirtualQueryEx
            return (T)Marshal.PtrToStructure(memAddress, typeof(T));
        }
    }

    public static class ComHelper
    {
        public static bool HRESULT_FAILED(int hr) => hr < 0;
    }

    // RestrictComInterfaceByAggregation is used to ensure that a wrapped COM object only responds to a specific interface
    // In particular, we don't want them to respond to IProvideClassInfo, which is broken in the VBE for some ITypeInfo implementations 
    public class RestrictComInterfaceByAggregation<T> : ICustomQueryInterface, IDisposable
    {
        private IntPtr _outerObject;
        private T _wrappedObject;

        public RestrictComInterfaceByAggregation(IntPtr outerObject, bool queryForType = true)
        {
            if (queryForType)
            {
                var ppv = IntPtr.Zero;
                var IID = typeof(T).GUID;
                if (ComHelper.HRESULT_FAILED(Marshal.QueryInterface(outerObject, ref IID, out _outerObject)))
                {
                    // allow null wrapping here
                    return;
                }
            }
            else
            {
                _outerObject = outerObject;
                Marshal.AddRef(_outerObject);
            }            

            var aggObjPtr = Marshal.CreateAggregatedObject(_outerObject, this);
            _wrappedObject = (T)Marshal.GetObjectForIUnknown(aggObjPtr);        // when this CCW object gets released, it will free the aggObjInner (well, after GC)
            Marshal.Release(aggObjPtr);         // _wrappedObject holds a reference to this now
        }

        public T WrappedObject { get => _wrappedObject; }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            if (_wrappedObject != null) Marshal.ReleaseComObject(_wrappedObject);
            if (_outerObject != IntPtr.Zero) Marshal.Release(_outerObject);
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;

            if (!_isDisposed)
            {
                if (iid == typeof(T).GUID)
                {
                    ppv = _outerObject;
                    Marshal.AddRef(_outerObject);
                    return CustomQueryInterfaceResult.Handled;
                }
            }
            
            return CustomQueryInterfaceResult.Failed;
        }
    }

    public class DisposableList<T> : IList<T>, IDisposable
        where T : IDisposable
    {
        private readonly IList<T> _list = new List<T>();

        public void Dispose() => ((IDisposable)this).Dispose();

        private bool _isDisposed;
        void IDisposable.Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            foreach (var element in _list)
            {
                element.Dispose();
            }
        }

        public IEnumerator<T> GetEnumerator() => _list.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public void Add(T item) => _list.Add(item);
        public void Clear() => _list.Clear();
        public bool Contains(T item) => _list.Contains(item);
        public void CopyTo(T[] array, int arrayIndex) => _list.CopyTo(array, arrayIndex);
        public bool Remove(T item) => _list.Remove(item);
        public int Count { get => _list.Count; }
        public bool IsReadOnly { get => _list.IsReadOnly; }
        
        public int IndexOf(T item) => _list.IndexOf(item);
        public void Insert(int index, T item) => _list.Insert(index, item);
        public void RemoveAt(int index) => _list.RemoveAt(index);
        public T this[int index]
        {
            get => _list[index];
            set => _list[index] = value;
        }
    }

    public class TypeInfoFunc : IDisposable
    {
        private TypeInfoWrapper _typeInfo;
        private ComTypes.FUNCDESC _funcDesc;
        private IntPtr _funcDescPtr;
        private string[] _names = new string[255];   // includes argument names
        private int _cNames = 0;

        public ComTypes.FUNCDESC FuncDesc { get => _funcDesc; }

        public TypeInfoFunc(TypeInfoWrapper typeInfo, int funcIndex)
        {
            _typeInfo = typeInfo;

            ((ComTypes.ITypeInfo)_typeInfo).GetFuncDesc(funcIndex, out _funcDescPtr);
            _funcDesc = StructHelper.ReadStructure<ComTypes.FUNCDESC>(_funcDescPtr);

            ((ComTypes.ITypeInfo)_typeInfo).GetNames(_funcDesc.memid, _names, _names.Length, out _cNames);
            if (_cNames == 0) _names[0] = "[unnamed]";
        }

        public void Dispose()
        {
            if (_funcDescPtr != IntPtr.Zero) ((ComTypes.ITypeInfo)_typeInfo).ReleaseFuncDesc(_funcDescPtr);
            _funcDescPtr = IntPtr.Zero;
        }

        public string Name { get => _names[0]; }
        public int ParamCount { get => _funcDesc.cParams;  }
        
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
                if (_funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF))
                {
                    return PROCKIND.PROCKIND_SET;
                }
                else if (_funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT))
                {
                    return PROCKIND.PROCKIND_LET;
                }
                else if (_funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYGET))
                {
                    return PROCKIND.PROCKIND_GET;
                }
                else
                {
                    return PROCKIND.PROCKIND_PROC;
                }
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

            output.AppendLine("- member: " + namesInfo + " [id 0x" + _funcDesc.memid.ToString("X") + ", " + _funcDesc.invkind + "]");
        }
    }

    public class TypeInfoReference
    {
        readonly string _rawString;
        readonly Guid _guid;
        readonly uint _majorVersion;
        readonly uint _minorVersion;
        readonly uint _lcid;
        readonly string _path;
        readonly string _name;

        public string RawString { get => _rawString; }
        public Guid GUID { get => _guid; }
        public uint MajorVersion { get => _majorVersion; }
        public uint MinorVersion { get => _minorVersion; }
        public uint LCID { get => _lcid; }
        public string Path { get => _path; }
        public string Name { get => _name; }

        public TypeInfoReference(string referenceStringRaw)
        {
            // Example: "*\G{000204EF-0000-0000-C000-000000000046}#4.1#9#C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7\VBE7.DLL#Visual Basic For Applications"
            // LibidReference defined at https://msdn.microsoft.com/en-us/library/dd922767(v=office.12).aspx
            // The string is split into 5 parts, delimited by #

            _rawString = referenceStringRaw;

            var referenceStringParts = referenceStringRaw.Split(new char[] { '#' }, 5);
            if (referenceStringParts.Length != 5)        
            {
                throw new ArgumentException($"Invalid reference string got {referenceStringRaw}.  Expected 5 parts.");
            }

            _guid = Guid.Parse(referenceStringParts[0].Substring(3));
            var versionSplit = referenceStringParts[1].Split(new char[] { '.' }, 2);
            if (versionSplit.Length != 2)
            {
                throw new ArgumentException($"Invalid reference string got {referenceStringRaw}.  Invalid version string.");
            }
            _majorVersion = uint.Parse(versionSplit[0], NumberStyles.AllowHexSpecifier);
            _minorVersion = uint.Parse(versionSplit[1], NumberStyles.AllowHexSpecifier);

            _lcid = uint.Parse(referenceStringParts[2], NumberStyles.AllowHexSpecifier);
            _path = referenceStringParts[3];
            _name = referenceStringParts[4];
        }

        public void Document(StringLineBuilder output)
        {
            output.AppendLine("- VBE Reference: " + Name + " [path: " + Path + ", majorVersion: " + MajorVersion +
                                ", minorVersion: " + MinorVersion + ", guid: " + GUID + ", lcid: " + LCID + "]");
        }
    }

    public class TypeInfoVar : IDisposable
    { 
        private TypeInfoWrapper _typeInfo;
        private ComTypes.VARDESC _varDesc;
        private IntPtr _varDescPtr;
        private string _name;

        public string Name { get => _name; }

        public TypeInfoVar(TypeInfoWrapper typeInfo, int index)
        {
            _typeInfo = typeInfo;

            ((ComTypes.ITypeInfo)_typeInfo).GetVarDesc(index, out _varDescPtr);
            _varDesc = StructHelper.ReadStructure<ComTypes.VARDESC>(_varDescPtr);
            
            int _cNames = 0;
            var _names = new string[1];
            if (_varDesc.memid != (int)TypeLibConsts.MEMBERID_NIL)
            {
                _cNames = 0;
                ((ComTypes.ITypeInfo)_typeInfo).GetNames(_varDesc.memid, _names, _names.Length, out _cNames);
                _name = _names[0];
            }
            else
            {
                _name = "{unknown}";     // VBA Constants appear in the typelib with no name
            }
        }

        public void Dispose()
        {
            if (_varDescPtr != IntPtr.Zero) ((ComTypes.ITypeInfo)_typeInfo).ReleaseVarDesc(_varDescPtr);
            _varDescPtr = IntPtr.Zero;
        }

        public void Document(StringLineBuilder output) 
        {
            output.AppendLine("- field: " + _name + " [id 0x" + _varDesc.memid.ToString("X") + "]");
        }
    }
    
    public class IIndexedCollection<TItem> : IEnumerable<TItem>
        where TItem : class
    {
        IEnumerator IEnumerable.GetEnumerator() => new IIndexedCollectionEnumerator<IIndexedCollection<TItem>, TItem>(this);
        public IEnumerator<TItem> GetEnumerator() => new IIndexedCollectionEnumerator<IIndexedCollection<TItem>, TItem>(this);

        virtual public int Count { get => throw new NotImplementedException(); }
        virtual public TItem GetItemByIndex(int index) => throw new NotImplementedException();
    }
    
    public class IIndexedCollectionEnumerator<TCollection, TItem> : IEnumerator<TItem>, IDisposable
        where TCollection : IIndexedCollection<TItem>
        where TItem : class
    {
        private TCollection _collection;
        private int _collectionCount;
        private int _index = -1;
        TItem _current;

        public IIndexedCollectionEnumerator(TCollection collection)
        {
            _collection = collection;
            _collectionCount = _collection.Count;       
        }

        public void Dispose()
        {
            // nothing to do here.
        }

        TItem IEnumerator<TItem>.Current => _current;
        object IEnumerator.Current => _current;

        public void Reset() => _index = -1;

        public bool MoveNext()
        {
            _current = default(TItem);
            _index++;
            if (_index >= _collectionCount) return false;
            _current = _collection.GetItemByIndex(_index);
            return true;
        }
    }

    // A wrapper for ITypeInfo provided by VBE, allowing safe managed consumption, plus adds StdModExecute functionality
    public class TypeInfoWrapper : ComTypes.ITypeInfo, IDisposable
    {
        private DisposableList<TypeInfoWrapper> _typeInfosWrapped;
        private TypeLibWrapper _containerTypeLib;
        public TypeLibWrapper Container { get => _containerTypeLib; }
        private int _containerTypeLibIndex;
        private bool _isUserFormBaseClass = false;
        private IntPtr _rawObjectPtr;
        private ComTypes.ITypeInfo _wrappedObjectRCW;

        private ComTypes.TYPEATTR _cachedAttributes;
        public ComTypes.TYPEATTR Attributes { get => _cachedAttributes; }

        private RestrictComInterfaceByAggregation<ComTypes.ITypeInfo> _ITypeInfo_Aggregator;
        private ComTypes.ITypeInfo target_ITypeInfo { get => _ITypeInfo_Aggregator?.WrappedObject ?? _wrappedObjectRCW; }

        private RestrictComInterfaceByAggregation<IVBEComponent> _IVBEComponent_Aggregator;
        private IVBEComponent target_IVBEComponent { get => _IVBEComponent_Aggregator?.WrappedObject; }

        private RestrictComInterfaceByAggregation<IVBETypeInfo> _IVBETypeInfo_Aggregator;
        private IVBETypeInfo target_IVBETypeInfo { get => _IVBETypeInfo_Aggregator?.WrappedObject; }

        public bool HasVBEExtensions { get => _IVBETypeInfo_Aggregator?.WrappedObject != null; }

        private bool _hasModuleScopeCompilationErrors;
        public bool HasModuleScopeCompilationErrors => _hasModuleScopeCompilationErrors;

        public class FuncsCollection : IIndexedCollection<TypeInfoFunc>
        {
            TypeInfoWrapper _parent;
            public FuncsCollection(TypeInfoWrapper parent) => _parent = parent;
            override public int Count { get => _parent.Attributes.cFuncs; }
            override public TypeInfoFunc GetItemByIndex(int index) => new TypeInfoFunc(_parent, index);
        }
        public FuncsCollection Funcs;

        public class VarsCollection : IIndexedCollection<TypeInfoVar>
        {
            TypeInfoWrapper _parent;
            public VarsCollection(TypeInfoWrapper parent) => _parent = parent;
            override public int Count { get => _parent.Attributes.cVars; }
            override public TypeInfoVar GetItemByIndex(int index) => new TypeInfoVar(_parent, index);
        }
        public VarsCollection Vars;

        public class ImplementedInterfacesCollection : IIndexedCollection<TypeInfoWrapper>
        {
            TypeInfoWrapper _parent;
            public ImplementedInterfacesCollection(TypeInfoWrapper parent) => _parent = parent;
            override public int Count { get => _parent.Attributes.cImplTypes; }
            override public TypeInfoWrapper GetItemByIndex(int index) => _parent.GetSafeImplementedTypeInfo(index);

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

            public bool DoesImplement(string interfaceProgId)
            {
                var progIdSplit = interfaceProgId.Split(new char[] { '.' }, 2);
                if (progIdSplit.Length != 2)
                {
                    throw new ArgumentException($"Expected a progid in the form of 'LibraryName.InterfaceName', got {interfaceProgId}");
                }
                return DoesImplement(progIdSplit[0], progIdSplit[1]);
            }

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
                _cachedAttributes = StructHelper.ReadStructure<ComTypes.TYPEATTR>(typeAttrPtr);
                target_ITypeInfo.ReleaseTypeAttr(typeAttrPtr);      // don't need to keep a hold of it, as _cachedAttributes is a copy
            }
            catch (Exception e)
            {
                if (e.HResult == (int)KnownComHResults.E_VBA_COMPILEERROR)
                {
                    _hasModuleScopeCompilationErrors = true;
                }

                // just mute the erorr and expose an empty type
                _cachedAttributes = new ComTypes.TYPEATTR();
            }

            // cache the container type library if it is available
            try
            {
                // We have to wrap the ITypeLib returned by GetContainingTypeLib
                // so we cast to our ITypeInfo_Ptrs interface in order to work with the raw IntPtrs
                IntPtr typeLibPtr = IntPtr.Zero;
                ((ITypeInfo_Ptrs)target_ITypeInfo).GetContainingTypeLib(out typeLibPtr, out _containerTypeLibIndex);
                _containerTypeLib = new TypeLibWrapper(typeLibPtr);  // takes ownership of the COM reference
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

        public TypeInfoWrapper(IntPtr rawObjectPtr, int? parentUserFormUniqueId = null)
        {
            _rawObjectPtr = rawObjectPtr;
            
            _ITypeInfo_Aggregator       = new RestrictComInterfaceByAggregation<ComTypes.ITypeInfo>(rawObjectPtr, queryForType: false);
            _IVBEComponent_Aggregator   = new RestrictComInterfaceByAggregation<IVBEComponent>(rawObjectPtr);
            _IVBETypeInfo_Aggregator    = new RestrictComInterfaceByAggregation<IVBETypeInfo>(rawObjectPtr);

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
            _IVBETypeInfo_Aggregator?.Dispose();
        }

        private TypeLibTextFields? _cachedTextFields;
        TypeLibTextFields CachedTextFields
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

        public string Name { get => CachedTextFields._name; }
        public string DocString { get => CachedTextFields._docString; }
        public int HelpContext { get => CachedTextFields._helpContext; }
        public string HelpFile { get => CachedTextFields._helpFile; }

        public string GetProgID() => (Container?.Name ?? "") + "." + CachedTextFields._name;

        public Guid GUID { get => Attributes.guid; }
        public TYPEKIND_VBE TypeKind { get => (TYPEKIND_VBE)Attributes.typekind; }
        
        public bool HasPredeclaredId { get => Attributes.wTypeFlags.HasFlag(ComTypes.TYPEFLAGS.TYPEFLAG_FPREDECLID); }
        public ComTypes.TYPEFLAGS Flags { get => Attributes.wTypeFlags; }

        private bool HasNoContainer() => _containerTypeLib == null;

        public bool CompileComponent()
        {
            if (HasVBEExtensions)
            {
                try
                {
                    target_IVBEComponent.CompileComponent();
                    return true;
                }
                catch (Exception e)
                {
                    if (e.HResult == (int)KnownComHResults.E_VBA_COMPILEERROR)
                    {
                        return false;
                    }
                    else
                    {
                        // this is more for debug purposes, as we can probably just return false in future.
                        throw new ArgumentException("Unrecognised VBE compiler error: \n" + e.ToString());
                    }
                }
            }
            else
            {
                throw new ArgumentException("This TypeInfo does not represent a VBE component, so we cannot compile it");
            }
        }

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

        // caller is responsible for calling ReleaseComObject
        public IDispatch GetStdModInstance()
        {
            if (HasVBEExtensions)
            {
                return target_IVBETypeInfo.GetStdModInstance();
            }
            else
            {
                throw new ArgumentException("This ITypeInfo is not hosted by the VBE, so does not support GetStdModInstance");
            }
        }

        public object StdModExecute(string name, Reflection.BindingFlags invokeAttr, object[] args = null)
        {
            if (HasVBEExtensions)
            {
                var StaticModule = GetStdModInstance();
                var retVal = StaticModule.GetType().InvokeMember(name, invokeAttr, null, StaticModule, args);
                Marshal.ReleaseComObject(StaticModule);
                return retVal;
            }
            else
            {
                throw new ArgumentException("This ITypeInfo is not hosted by the VBE, so does not support StdModExecute");
            }
        }

        // Gets the control ITypeInfo by looking for the corresponding getter on the form interface and returning its retval type
        // Supports UserForms.  what about Access forms etc
        public TypeInfoWrapper GetControlType(string controlName)
        {
            // TODO should encapsulate handling of raw datatypes
            foreach (var func in Funcs)
            {
                using (func)
                {
                    if ((func.Name == controlName) &&
                        (func.ProcKind == TypeInfoFunc.PROCKIND.PROCKIND_GET) &&
                        (func.ParamCount == 0) &&
                        (func.FuncDesc.elemdescFunc.tdesc.vt == (short)VarEnum.VT_PTR))
                    {
                        var retValElement = StructHelper.ReadStructure<ComTypes.ELEMDESC>(func.FuncDesc.elemdescFunc.tdesc.lpValue);
                        if (retValElement.tdesc.vt == (short)VarEnum.VT_USERDEFINED)
                        {
                            return GetSafeRefTypeInfo((int)retValElement.tdesc.lpValue);
                        }
                    }
                }
            }

            throw new ArgumentException($"TypeInfoWrapper::GetControlType failed. '{controlName}' control not found.");
        }

        public TypeInfoWrapper GetSafeRefTypeInfo(int hRef)
        {
            IntPtr typeInfoPtr = IntPtr.Zero;
            // we cast to our ITypeInfo_Ptrs interface in order to work with the raw IntPtr for aggregation
            ((ITypeInfo_Ptrs)target_ITypeInfo).GetRefTypeInfo(hRef, out typeInfoPtr);
            var outVal = new TypeInfoWrapper(typeInfoPtr, _isUserFormBaseClass ? (int?)hRef : null); // takes ownership of the COM reference
            _typeInfosWrapped = _typeInfosWrapped ?? new DisposableList<TypeInfoWrapper>();
            _typeInfosWrapped.Add(outVal);
            return outVal;
        }

        public TypeInfoWrapper GetSafeImplementedTypeInfo(int index)
        {
            target_ITypeInfo.GetRefTypeOfImplType(index, out int href);
            return GetSafeRefTypeInfo(href);
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
            foreach (var var in Vars)
            {
                using (var)
                {
                    var.Document(output);
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
    
    public struct TypeLibTextFields
    {
        public string _name;
        public string _docString;
        public int _helpContext;
        public string _helpFile;
    }

    // A wrapper for ITypeLib that exposes VBE ITypeInfos safely for managed consumption, plus adds ConditionalCompilationArguments property
    public class TypeLibWrapper : ComTypes.ITypeLib, IDisposable
    {
        private DisposableList<TypeInfoWrapper> _typeInfosWrapped;
        private readonly bool _wrappedObjectIsWeakReference;

        public class TypeInfosCollection : IIndexedCollection<TypeInfoWrapper>
        {
            TypeLibWrapper _parent;
            public TypeInfosCollection(TypeLibWrapper parent) => _parent = parent;
            override public int Count { get => _parent.TypesCount; }
            override public TypeInfoWrapper GetItemByIndex(int index) => _parent.GetSafeTypeInfoByIndex(index);

            public TypeInfoWrapper Find(string searchTypeName)
            {
                foreach (var typeInfo in this)
                {
                    if (typeInfo.Name == searchTypeName) return typeInfo;
                    typeInfo.Dispose();
                }
                return null;
            }

            public TypeInfoWrapper Get(string searchTypeName)
            {
                var retVal = Find(searchTypeName);
                if (retVal == null)
                {
                    throw new ArgumentException($"TypeInfosCollection::Get failed. '{searchTypeName}' component not found.");
                }
                return retVal;
            }
        }
        public TypeInfosCollection TypeInfos;
        
        public class ReferencesCollection : IIndexedCollection<TypeInfoReference>
        {
            TypeLibWrapper _parent;
            public ReferencesCollection(TypeLibWrapper parent) => _parent = parent;
            override public int Count { get => _parent.GetVBEReferencesCount(); }
            override public TypeInfoReference GetItemByIndex(int index) => _parent.GetVBEReferenceByIndex(index);
        }
        public ReferencesCollection VBEReferences;

        private TypeLibTextFields? _cachedTextFields;
        TypeLibTextFields CachedTextFields
        {
            get
            {
                if (!_cachedTextFields.HasValue)
                {
                    var cache = new TypeLibTextFields();
                    target_ITypeLib.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out cache._name, out cache._docString, out cache._helpContext, out cache._helpFile);
                    _cachedTextFields = cache;
                }
                return _cachedTextFields.Value; 
            }
        }

        public string Name      { get => CachedTextFields._name; }
        public string DocString { get => CachedTextFields._docString; }
        public int HelpContext  { get => CachedTextFields._helpContext; }
        public string HelpFile  { get => CachedTextFields._helpFile; }

        private ComTypes.ITypeLib target_ITypeLib;
        private IVBEProject target_IVBEProject;

        public bool HasVBEExtensions { get => target_IVBEProject != null; }

        public int GetVBEReferencesCount()
        {
            if (HasVBEExtensions)
            {
                return target_IVBEProject.GetReferencesCount();
            }
            else
            {
                throw new ArgumentException("This TypeLib does not represent a VBE project, so we cannot get reference strings from it");
            }
        }

        public TypeInfoReference GetVBEReferenceByIndex(int index)
        {
            if (HasVBEExtensions)
            {
                if (index < target_IVBEProject.GetReferencesCount())
                {
                    return new TypeInfoReference(target_IVBEProject.GetReferenceString(index));
                }

                throw new ArgumentException($"Specified index not valid for the references collection {index}.");
            }
            else
            {
                throw new ArgumentException("This TypeLib does not represent a VBE project, so we cannot get reference strings from it");
            }
        }

        public TypeInfoReference GetVBEReferenceByGuid(Guid referenceGuid)
        {
            if (HasVBEExtensions)
            {
                foreach (var reference in VBEReferences)
                {
                    if (reference.GUID == referenceGuid)
                    {
                        return reference;
                    }
                }

                throw new ArgumentException($"Specified GUID not found in references collection {referenceGuid}.");
            }
            else
            {
                throw new ArgumentException("This TypeLib does not represent a VBE project, so we cannot get reference strings from it");
            }
        }

        

        /*
         This is not yet used, but here in case we want to use this interface at some point.
        private RestrictComInterfaceByAggregation<IVBEProject2> _cachedIVBEProject2;   NEEDS DISPOSING
        private IVBEProjectEx2 target_IVBEProject2
        {
            get
            {
                if (_cachedIVBProjectEx2 == null)
                {
                    if (HasVBEExtensions)
                    {
                        // This internal VBE interface doesn't have a queryable IID.  
                        // The vtable for this interface directly preceeds the _IVBProjectEx, and we can access it through an aggregation helper
                        var objIVBProjectExPtr = Marshal.GetComInterfaceForObject(_wrappedObject, typeof(IVBEProject));
                        _cachedIVBProjectEx2 = new RestrictComInterfaceByAggregation<IVBEProject2>(objIVBProjectExPtr - IntPtr.Size, queryForType: false);
                    }
                    else
                    {
                        throw new ArgumentException("This ITypeLib is not hosted by the VBE, so does not support _IVBProjectEx");
                    }
                }

                return (IVBEProject2)_cachedIVBProjectEx2.WrappedObject;
            }
        }*/

        public static TypeLibWrapper FromVBProject(IVBProject vbProject)
        {
            using (var references = vbProject.References)
            {
                // Now we've got the references object, we can read the internal object structure to grab the ITypeLib
                var internalReferencesObj = StructHelper.ReadComObjectStructure<VBEReferencesObj>(references.Target);

                return new TypeLibWrapper(internalReferencesObj.TypeLib);
            }
        }

        private void InitCommon()
        {
            TypeInfos = new TypeInfosCollection(this);
            target_IVBEProject = target_ITypeLib as IVBEProject;
            if (HasVBEExtensions) VBEReferences = new ReferencesCollection(this);
        }

        public TypeLibWrapper(IntPtr rawObjectPtr)
        {
            target_ITypeLib = (ComTypes.ITypeLib)Marshal.GetObjectForIUnknown(rawObjectPtr);
            Marshal.Release(rawObjectPtr);         // _wrappedObject holds a reference to this now
            InitCommon();
        }

        public TypeLibWrapper(ComTypes.ITypeLib rawTypeInfo)
        {
            target_ITypeLib = rawTypeInfo;
            _wrappedObjectIsWeakReference = true;
            InitCommon();
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            _typeInfosWrapped?.Dispose();
            if (!_wrappedObjectIsWeakReference) Marshal.ReleaseComObject(target_ITypeLib);
        }

        public TypeInfoWrapper GetSafeTypeInfoByIndex(int index)
        {
            IntPtr typeInfoPtr = IntPtr.Zero;
            // We cast to our IVBETypeLib interface in order to work with the raw IntPtr for aggregation
            ((ITypeLib_Ptrs)target_ITypeLib).GetTypeInfo(index, out typeInfoPtr);
            var outVal = new TypeInfoWrapper(typeInfoPtr);
            _typeInfosWrapped = _typeInfosWrapped ?? new DisposableList<TypeInfoWrapper>();
            _typeInfosWrapped.Add(outVal);
            return outVal;
        }

        public int TypesCount
        {
            get => target_ITypeLib.GetTypeInfoCount();
        }

        private ComTypes.TYPELIBATTR? _cachedLibAttribs;
        public ComTypes.TYPELIBATTR Attributes
        {
            get
            {
                if (!_cachedLibAttribs.HasValue)
                {
                    target_ITypeLib.GetLibAttr(out IntPtr typeLibAttributesPtr);
                    _cachedLibAttribs = StructHelper.ReadStructure<ComTypes.TYPELIBATTR>(typeLibAttributesPtr);
                    target_ITypeLib.ReleaseTLibAttr(typeLibAttributesPtr);          // no need to keep open.  copied above
                }
                return _cachedLibAttribs.Value;
            }
        }

        public bool CompileProject()
        {
            if (HasVBEExtensions)
            {
                try
                {
                    target_IVBEProject.CompileProject();
                    return true;
                }
                catch (Exception e)
                {
                    if (e.HResult == (int)KnownComHResults.E_VBA_COMPILEERROR)
                    {
                        return false;
                    }
                    else
                    {
                        // this is more for debug purposes, as we can probably just return false in future.
                        throw new ArgumentException("Unrecognised VBE compiler error: \n" + e.ToString());
                    }
                }
            }
            else
            {
                throw new ArgumentException("This TypeLib does not represent a VBE project, so we cannot compile it");
            }
        }

        public string ConditionalCompilationArgumentsRaw
        {
            get
            {
                if (HasVBEExtensions)
                {
                    return target_IVBEProject.GetConditionalCompilationArgs();
                }
                else
                {
                    throw new ArgumentException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }
            }

            set
            {
                if (HasVBEExtensions)
                {
                    target_IVBEProject.SetConditionalCompilationArgs(value);
                }
                else
                {
                    throw new ArgumentException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }
            }
        }

        public Dictionary<string, string> ConditionalCompilationArguments
        {
            get
            {
                if (HasVBEExtensions)
                {
                    string args = target_IVBEProject.GetConditionalCompilationArgs();

                    if (args.Length > 0)
                    {
                        string[] argsArray = args.Split(new[] { ':' });
                        return argsArray.Select(item => item.Split('=')).ToDictionary(s => s[0], s => s[1]);
                    }
                    else
                    {
                        return new Dictionary<string, string>();
                    }
                }
                else
                {
                    throw new ArgumentException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }
            }

            set
            {
                if (HasVBEExtensions)
                {
                    var rawArgsString = string.Join(" : ", value.Select(x => x.Key + " = " + x.Value));
                    ConditionalCompilationArgumentsRaw = rawArgsString;
                }
                else
                {
                    throw new ArgumentException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }
            }
        }
        
        public void Document(StringLineBuilder output)
        {
            output.AppendLine();
            output.AppendLine("================================================================================");
            output.AppendLine();

            var libName = Name ?? "[VBA.Immediate.Window]";     

            output.AppendLine("ITypeLib: " + Name);
            output.AppendLineNoNullChars("- Documentation: " + DocString);
            output.AppendLineNoNullChars("- HelpContext: " + HelpContext);
            output.AppendLineNoNullChars("- HelpFile: " + HelpFile);
            output.AppendLine("- Guid: " + Attributes.guid);
            output.AppendLine("- Lcid: " + Attributes.lcid);
            output.AppendLine("- SysKind: " + Attributes.syskind);
            output.AppendLine("- LibFlags: " + Attributes.wLibFlags);
            output.AppendLine("- MajorVer: " + Attributes.wMajorVerNum);
            output.AppendLine("- MinorVer: " + Attributes.wMinorVerNum);
            output.AppendLine("- HasVBEExtensions: " + HasVBEExtensions);
            if (HasVBEExtensions)
            {
                output.AppendLine("- VBE Conditional Compilation Arguments: " + ConditionalCompilationArguments);

                foreach (var reference in VBEReferences)
                {
                    reference.Document(output);
                }
            }

            output.AppendLine("- TypeCount: " + TypesCount);

            foreach (var typeInfo in TypeInfos)
            {
                using (typeInfo)
                {
                    typeInfo.Document(output, libName, 0);
                }
            }
        }

        // We have to wrap the ITypeInfo returned by GetTypeInfoOfGuid
        // so we cast to our IVBETypeLib interface in order to work with the raw IntPtr for aggregation
        void ComTypes.ITypeLib.GetTypeInfoOfGuid(ref Guid guid, out ComTypes.ITypeInfo ppTInfo)
        {
            IntPtr typeInfoPtr = IntPtr.Zero;
            ((ITypeLib_Ptrs)target_ITypeLib).GetTypeInfoOfGuid(guid, out typeInfoPtr);
            var outVal = new TypeInfoWrapper(typeInfoPtr);  // takes ownership of the COM reference
            ppTInfo = outVal;

            _typeInfosWrapped = _typeInfosWrapped ?? new DisposableList<TypeInfoWrapper>();
            _typeInfosWrapped.Add(outVal);
        }
        void ComTypes.ITypeLib.GetTypeInfo(int index, out ComTypes.ITypeInfo ppTI)
            => ppTI = GetSafeTypeInfoByIndex(index);   // We have to wrap the ITypeInfo returned by GetTypeInfo
        int ComTypes.ITypeLib.GetTypeInfoCount()
            => target_ITypeLib.GetTypeInfoCount();
        void ComTypes.ITypeLib.GetTypeInfoType(int index, out ComTypes.TYPEKIND pTKind)
            => target_ITypeLib.GetTypeInfoType(index, out pTKind);
        void ComTypes.ITypeLib.GetLibAttr(out IntPtr ppTLibAttr)
            => target_ITypeLib.GetLibAttr(out ppTLibAttr);
        void ComTypes.ITypeLib.GetTypeComp(out ComTypes.ITypeComp ppTComp)
            => target_ITypeLib.GetTypeComp(out ppTComp);
        void ComTypes.ITypeLib.GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile)
            => target_ITypeLib.GetDocumentation(index, out strName, out strDocString, out dwHelpContext, out strHelpFile);
        bool ComTypes.ITypeLib.IsName(string szNameBuf, int lHashVal)
            => target_ITypeLib.IsName(szNameBuf, lHashVal);
        void ComTypes.ITypeLib.FindName(string szNameBuf, int lHashVal, ComTypes.ITypeInfo[] ppTInfo, int[] rgMemId, ref short pcFound)
            => target_ITypeLib.FindName(szNameBuf, lHashVal, ppTInfo, rgMemId, pcFound);
        void ComTypes.ITypeLib.ReleaseTLibAttr(IntPtr pTLibAttr)
            => target_ITypeLib.ReleaseTLibAttr(pTLibAttr);
    }

    // class for iterating over the double linked list of ITypeLibs provided by the VBE
    public class VBETypeLibsIterator : IEnumerable<TypeLibWrapper>, IEnumerator<TypeLibWrapper>, IDisposable
    {
        private IntPtr _currentTypeLibPtr;
        private VBETypeLibObj _currentTypeLibStruct;
        private bool _isStart;

        public VBETypeLibsIterator(IntPtr typeLibPtr)
        {
            _currentTypeLibPtr = typeLibPtr;
            _currentTypeLibStruct = StructHelper.ReadStructureSafe<VBETypeLibObj>(_currentTypeLibPtr);
            Reset();
        }

        public void Dispose()
        {
            // nothing to do here, we don't own anything that needs releasing
        }

        IEnumerator IEnumerable.GetEnumerator() => this;
        public IEnumerator<TypeLibWrapper> GetEnumerator() => this;

        public IntPtr GetCurrentReference()
        {
            Marshal.AddRef(_currentTypeLibPtr);
            return _currentTypeLibPtr;
        }

        TypeLibWrapper IEnumerator<TypeLibWrapper>.Current => new TypeLibWrapper(GetCurrentReference());
        object IEnumerator.Current => new TypeLibWrapper(GetCurrentReference());

        public void Reset()  // walk back to the first project in the chain
        {
            while (_currentTypeLibStruct.Prev != IntPtr.Zero)
            {
                _currentTypeLibPtr = _currentTypeLibStruct.Prev;
                _currentTypeLibStruct = StructHelper.ReadStructureSafe<VBETypeLibObj>(_currentTypeLibPtr);
            }
            _isStart = true;
        }

        public bool MoveNext()
        {
            if (_isStart)
            {
                _isStart = false;  // MoveNext is called before accessing the very first item
                return true;
            }

            if (_currentTypeLibStruct.Next == IntPtr.Zero) return false;

            _currentTypeLibPtr = _currentTypeLibStruct.Next;
            _currentTypeLibStruct = StructHelper.ReadStructureSafe<VBETypeLibObj>(_currentTypeLibPtr);
            return true;
        }
    }

    // the main class for hooking into the live ITypeLibs provided by the VBE
    public class VBETypeLibsAccessor : DisposableList<TypeLibWrapper>, IDisposable
    {
        public VBETypeLibsAccessor(IVBE ide)
        {
            // We need at least one project in the VBE.VBProjects collection to be accessible (i.e. unprotected)
            // in order to get access to the list of loaded project TypeLibs using this method

            foreach (var project in ide.VBProjects)
            {
                using (project)
                {
                    try
                    {
                        using (var references = project.References)
                        {
                            // Now we've got the references object, we can read the internal object structure to grab the ITypeLib
                            var internalReferencesObj = StructHelper.ReadComObjectStructure<VBEReferencesObj>(references.Target);

                            // Now we've got this one internalReferencesObj.typeLib, we can iterate through ALL loaded project TypeLibs
                            using (var typeLibIterator = new VBETypeLibsIterator(internalReferencesObj.TypeLib))
                            {
                                foreach (var typeLib in typeLibIterator)
                                {
                                    Add(typeLib);
                                }
                            }
                        }

                        // we only need access to a single VBProject References object to make it work, so we can return now.
                        return;
                    }
                    catch (Exception)
                    {
                        // probably a protected project, just move on to the next project.
                    }
                }
            }

            // return an empty list on error
        }

        public TypeLibWrapper Find(string searchLibName)
        {
            foreach (var typeLib in this)
            {
                if (typeLib.Name == searchLibName)
                {
                    return typeLib;
                }
            }
            return null;
        }

        public TypeLibWrapper Get(string searchLibName)
        {
            var retVal = Find(searchLibName);
            if (retVal == null)
            {
                throw new ArgumentException($"TypeLibWrapper::Get failed. '{searchLibName}' component not found.");
            }
            return retVal;
        }
    }
}
