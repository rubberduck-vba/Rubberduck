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

// TODO add memory address validation in ReadStructureSafe
// TODO split into TypeInfos.cs
// references expose the raw ITypeLibs

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
    /// Extension to StringBuilder to allow adding text line by line.
    /// </summary>
    public class StringLineBuilder
    {
        StringBuilder _document = new StringBuilder();

        public override string ToString() => _document.ToString();

        public void AppendLine(string value = "")
            => _document.Append(value + "\r\n");

        public void AppendLineNoNullChars(string value)
            => AppendLine(value.Replace("\0", string.Empty));
    }

    /// <summary>
    /// Encapsulates reading unmanaged memory into managed structures
    /// </summary>
    public static class StructHelper
    {
        /// <summary>
        /// Takes a COM object, and reads the unmanaged memory given by its pointer, allowing us to read internal fields
        /// </summary>
        /// <typeparam name="T">the type of structure to return</typeparam>
        /// <param name="comObj">the COM object</param>
        /// <returns>the requested structure T</returns>
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

        /// <summary>
        /// Takes an unmanaged memory address and reads the unmanaged memory given by its pointer
        /// </summary>
        /// <typeparam name="T">the type of structure to return</typeparam>
        /// <param name="memAddress">the unamanaged memory address to read</param>
        /// <returns>the requested structure T</returns>
        public static T ReadStructure<T>(IntPtr memAddress)
        {
            if (memAddress == IntPtr.Zero) return default(T);
            return (T)Marshal.PtrToStructure(memAddress, typeof(T));
        }

        /// <summary>
        /// Takes an unmanaged memory address and reads the unmanaged memory given by its pointer, 
        /// with memory address validation for added protection
        /// </summary>
        /// <typeparam name="T">the type of structure to return</typeparam>
        /// <param name="memAddress">the unamanaged memory address to read</param>
        /// <returns>the requested structure T</returns>
        public static T ReadStructureSafe<T>(IntPtr memAddress)
        {
            if (memAddress == IntPtr.Zero) return default(T);
            return (T)Marshal.PtrToStructure(memAddress, typeof(T));
        }
    }

    /// <summary>
    /// Ensures that a wrapped COM object only responds to a specific COM interface.
    /// </summary>
    /// <typeparam name="T">The COM interface for restriction</typeparam>
    public class RestrictComInterfaceByAggregation<T> : ICustomQueryInterface, IDisposable
    {
        private IntPtr _outerObject;
        private T _wrappedObject;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="outerObject">The object that needs interface requests filtered</param>
        /// <param name="queryForType">determines whether we call QueryInterface for the interface or not</param>
        /// <remarks>if the passed in outerObject is known to point to the correct vtable for the interface, then queryForType can be false</remarks>
        /// <returns>if outerObject is IntPtr.Zero, then a null wrapper, else an aggregated wrapper</returns>
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

    /// <summary>
    /// A disposable list that encapsulates the disposing of its elements
    /// </summary>
    /// <typeparam name="T"></typeparam>
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
    
    /// <summary>
    /// A class that represents a function definition within a typeinfo
    /// </summary>
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

    /// <summary>
    /// A class that represents a reference within a VBA project
    /// </summary>
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

    /// <summary>
    /// A class that represents a variable or field within a typeinfo
    /// </summary>
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
    
    /// <summary>
    /// A base class for exposing an enumerable collection through an index based accessor
    /// </summary>
    /// <typeparam name="TItem">the collection element type</typeparam>
    public abstract class IIndexedCollectionBase<TItem> : IEnumerable<TItem>
        where TItem : class
    {
        IEnumerator IEnumerable.GetEnumerator() => new IIndexedCollectionEnumerator<IIndexedCollectionBase<TItem>, TItem>(this);
        public IEnumerator<TItem> GetEnumerator() => new IIndexedCollectionEnumerator<IIndexedCollectionBase<TItem>, TItem>(this);

        abstract public int Count { get; }
        abstract public TItem GetItemByIndex(int index);
    }

    /// <summary>
    /// The enumerator implementation for IIndexedCollectionBase
    /// </summary>
    /// <typeparam name="TCollection">the IIndexedCollectionBase<> type</typeparam>
    /// <typeparam name="TItem">the collection element type</typeparam>
    public class IIndexedCollectionEnumerator<TCollection, TItem> : IEnumerator<TItem>
        where TCollection : IIndexedCollectionBase<TItem>
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

        /// <summary>
        /// Exposes an enumerable collection of functions provided by the ITypeInfo
        /// </summary>
        public class FuncsCollection : IIndexedCollectionBase<TypeInfoFunc>
        {
            TypeInfoWrapper _parent;
            public FuncsCollection(TypeInfoWrapper parent) => _parent = parent;
            override public int Count { get => _parent.Attributes.cFuncs; }
            override public TypeInfoFunc GetItemByIndex(int index) => new TypeInfoFunc(_parent, index);

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
            TypeInfoWrapper _parent;
            public VarsCollection(TypeInfoWrapper parent) => _parent = parent;
            override public int Count { get => _parent.Attributes.cVars; }
            override public TypeInfoVar GetItemByIndex(int index) => new TypeInfoVar(_parent, index);
        }
        public VarsCollection Vars;

        /// <summary>
        /// Exposes an enumerable collection of implemented interfaces provided by the ITypeInfo
        /// </summary>
        public class ImplementedInterfacesCollection : IIndexedCollectionBase<TypeInfoWrapper>
        {
            TypeInfoWrapper _parent;
            public ImplementedInterfacesCollection(TypeInfoWrapper parent) => _parent = parent;
            override public int Count { get => _parent.Attributes.cImplTypes; }
            override public TypeInfoWrapper GetItemByIndex(int index) => _parent.GetSafeImplementedTypeInfo(index);

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
                _cachedAttributes = StructHelper.ReadStructure<ComTypes.TYPEATTR>(typeAttrPtr);
                target_ITypeInfo.ReleaseTypeAttr(typeAttrPtr);      // don't need to keep a hold of it, as _cachedAttributes is a copy
            }
            catch (Exception e)
            {
                // If there is a compilation error outside of a procedure code block, the type information is not available for that component.
                // We detect this, via the E_VBA_COMPILEERROR error 
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

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rawObjectPtr">The raw unmanaged pointer to the ITypeInfo</param>
        /// <param name="parentUserFormUniqueId">used internally for providing a name for UserForm base classes</param>
        public TypeInfoWrapper(IntPtr rawObjectPtr, int? parentUserFormUniqueId = null)
        {
            _rawObjectPtr = rawObjectPtr;

            // We have to restrict interface requests to VBE hosted ITypeInfos due to a bug in their implementation.
            // See TypeInfoWrapper class XML doc for details.

            // queryForType is passed as false for ITypeInfo here, as rawObjectPtr is known to point to the ITypeInfo vtable
            // additionally allowing it to query for ITypeInfo gives a _different_, and more prohibitive implementation of ITypeInfo
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

        public string GetProgID() => (Container?.Name ?? "{unnamedlibrary}") + "." + CachedTextFields._name;

        public Guid GUID { get => Attributes.guid; }
        public TYPEKIND_VBE TypeKind { get => (TYPEKIND_VBE)Attributes.typekind; }
        
        public bool HasPredeclaredId { get => Attributes.wTypeFlags.HasFlag(ComTypes.TYPEFLAGS.TYPEFLAG_FPREDECLID); }
        public ComTypes.TYPEFLAGS Flags { get => Attributes.wTypeFlags; }

        private bool HasNoContainer() => _containerTypeLib == null;

        /// <summary>
        /// Silently compiles the individual VBE component (class/module etc)
        /// </summary>
        /// <returns>true if this module, plus any direct dependent modules compile successfully</returns>
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
            if (HasVBEExtensions)
            {
                return target_IVBETypeInfo.GetStdModAccessor();
            }
            else
            {
                throw new ArgumentException("This ITypeInfo is not hosted by the VBE, so does not support GetStdModAccessor");
            }
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
            if (HasVBEExtensions)
            {
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
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    Marshal.ReleaseComObject(staticModule);
                }
            }
            else
            {
                throw new ArgumentException("This ITypeInfo is not hosted by the VBE, so does not support StdModExecute");
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
                        var retValElement = StructHelper.ReadStructure<ComTypes.ELEMDESC>(func.FuncDesc.elemdescFunc.tdesc.lpValue);
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
                        var retValElementOuterPtr = StructHelper.ReadStructure<ComTypes.ELEMDESC>(func.FuncDesc.lprgelemdescParam);
                        if (retValElementOuterPtr.tdesc.vt == (short)VarEnum.VT_PTR)
                        {
                            var retValElementInnerPtr = StructHelper.ReadStructure<ComTypes.ELEMDESC>(retValElementOuterPtr.tdesc.lpValue);
                            if (retValElementInnerPtr.tdesc.vt == (short)VarEnum.VT_PTR)
                            {
                                var retValElement = StructHelper.ReadStructure<ComTypes.ELEMDESC>(retValElementInnerPtr.tdesc.lpValue);

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

    /// <summary>
    /// A wrapper for ITypeLib objects, with specific extensions for VBE hosted ITypeLibs
    /// </summary>
    /// <remarks>
    /// allow safe managed consumption, plus adds ConditionalCompilationArguments property, 
    /// VBEReferences collection, and CompileProject method.
    /// Can also be cast to ComTypes.ITypeLib for raw access to the underlying type library
    /// </remarks>
    public class TypeLibWrapper : ComTypes.ITypeLib, IDisposable
    {
        private DisposableList<TypeInfoWrapper> _typeInfosWrapped;
        private readonly bool _wrappedObjectIsWeakReference;

        /// <summary>
        /// Exposes an enumerable collection of TypeInfo objects exposed by this ITypeLib
        /// </summary>
        public class TypeInfosCollection : IIndexedCollectionBase<TypeInfoWrapper>
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
        
        /// <summary>
        /// Exposes an enumerable collection of references used by the VBE type library
        /// </summary>
        public class ReferencesCollection : IIndexedCollectionBase<TypeInfoReference>
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

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rawObjectPtr">The raw unamanaged ITypeLib pointer</param>
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

        /// <summary>
        /// Silently compiles the whole VBA project represented by this ITypeLib
        /// </summary>
        /// <returns>true if the compilation succeeds</returns>
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

        /// <summary>
        /// Exposes the raw conditional compilation arguments defined in the BA project represented by this ITypeLib
        /// format:  "foo = 1 : bar = 2"
        /// </summary>
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

        /// <summary>
        /// Exposes the conditional compilation arguments defined in the BA project represented by this ITypeLib
        /// as a dictionary of key/value pairs
        /// </summary>
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

    /// <summary>
    /// An enumerable class for iterating over the double linked list of ITypeLibs provided by the VBE 
    /// </summary>
    public class VBETypeLibsIterator : IEnumerable<TypeLibWrapper>, IEnumerator<TypeLibWrapper>
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

    /// <summary>
    /// The root class for hooking into the live ITypeLibs provided by the VBE
    /// </summary>
    /// <remarks>
    /// WARNING: when using VBETypeLibsAccessor directly, do not cache it
    ///   The VBE provides LIVE type library information, so consider it a snapshot at that very moment when you are dealing with it
    ///   Make sure you call VBETypeLibsAccessor.Dispose() as soon as you have done what you need to do with it.
    ///   Once control returns back to the VBE, you must assume that all the ITypeLib/ITypeInfo pointers are now invalid.
    /// </remarks>
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
