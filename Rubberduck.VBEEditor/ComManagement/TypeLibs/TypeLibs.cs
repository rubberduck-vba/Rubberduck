using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibsAbstract;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Reflection = System.Reflection;

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
    public class StructHelper
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

    // RestrictComInterfaceByAggregation is used to ensure that a wrapped COM object only responds to a specific interface
    // In particular, we don't want them to respond to IProvideClassInfo, which is broken in the VBE for some ITypeInfo implementations 
    public class RestrictComInterfaceByAggregation<T> : ICustomQueryInterface, IDisposable
    {
        public IntPtr _outerObject;
        private T _wrappedObject;

        public RestrictComInterfaceByAggregation(IntPtr outerObject, bool queryForType = true)
        {
            if (queryForType)
            {
                var ppv = IntPtr.Zero;
                var IID = typeof(T).GUID;
                if (Marshal.QueryInterface(outerObject, ref IID, out _outerObject) < 0)
                {
                    // allow a null wrapper here
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

    // A wrapper for ITypeInfo provided by VBE, allowing safe managed consumption, plus adds StdModExecute functionality
    public class TypeInfoWrapper : ComTypes.ITypeInfo, IDisposable
    {
        private DisposableList<TypeInfoWrapper> _typeInfosWrapped;
        private TypeLibWrapper _containerTypeLib;
        private int _containerTypeLibIndex;
        private bool _isUserFormBaseClass = false;
        private ComTypes.TYPEATTR _cachedAttributes;
        private IntPtr _rawObjectPtr;

        private string _name;
        private string _docString;
        private int _helpContext;
        private string _helpFile;

        public string Name { get => _name; }
        public string DocString { get => _docString; }
        public int HelpContext { get => _helpContext; }
        public string HelpFile { get => _helpFile; }

        private ComTypes.ITypeInfo _wrappedObjectRCW;

        private RestrictComInterfaceByAggregation<ComTypes.ITypeInfo>   _ITypeInfo_Aggregator;
        private ComTypes.ITypeInfo  target_ITypeInfo        { get => _ITypeInfo_Aggregator?.WrappedObject ?? _wrappedObjectRCW; }
        
        private RestrictComInterfaceByAggregation<IVBEComponent>        _IVBEComponent_Aggregator;
        private IVBEComponent       target_IVBEComponent    { get => _IVBEComponent_Aggregator?.WrappedObject; }

        private RestrictComInterfaceByAggregation<IVBETypeInfo>         _IVBETypeInfo_Aggregator;
        private IVBETypeInfo        target_IVBETypeInfo     { get => _IVBETypeInfo_Aggregator?.WrappedObject; }

        public bool HasVBEExtensions() => _IVBETypeInfo_Aggregator?.WrappedObject != null;

        private bool _hasModuleScopeCompilationErrors;
        public bool HasModuleScopeCompilationErrors => _hasModuleScopeCompilationErrors;

        private void InitCommon()
        {
            IntPtr typeAttrPtr = IntPtr.Zero;
            try
            {
                GetTypeAttr(out typeAttrPtr);
                _cachedAttributes = StructHelper.ReadStructure<ComTypes.TYPEATTR>(typeAttrPtr);
                ReleaseTypeAttr(typeAttrPtr);      // don't need to keep a hold of it, as _cachedAttributes is a copy
            }
            catch (Exception e)
            {
                if (e.HResult == (int)VBECompilerConsts.E_VBA_COMPILEERROR)
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
            catch (Exception)
            {
                // it is acceptable for a type to not have a container, as types can be runtime generated.
            }

            if (Name == null) target_ITypeInfo.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out _name, out _docString, out _helpContext, out _helpFile);
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
                _name = "_UserFormBase{unnamed}#" + parentUserFormUniqueId;
            }

            InitCommon();
            DetectUserFormClass();
        }

        public bool HasPredeclaredId { get => _cachedAttributes.wTypeFlags.HasFlag(ComTypes.TYPEFLAGS.TYPEFLAG_FPREDECLID); }

        private bool HasNoContainer() => _containerTypeLib == null;

        public bool CompileComponent()
        {
            if (HasVBEExtensions())
            {
                try
                {
                    target_IVBEComponent.CompileComponent();
                    return true;
                }
                catch (Exception e)
                {
                    if (e.HResult == (int)VBECompilerConsts.E_VBA_COMPILEERROR)
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
            if ((_cachedAttributes.typekind == ComTypes.TYPEKIND.TKIND_COCLASS) &&
                    HasNoContainer() &&
                    (_cachedAttributes.cImplTypes == 2) && 
                    (Name == "Form"))
            {
                // we can be 99.999999% sure it IS the runtime generated UserForm base class
                _isUserFormBaseClass = true;
            }
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

        // We have to wrap the ITypeInfo returned by GetRefTypeInfo
        // so we cast to our ITypeInfo_Ptrs interface in order to work with the raw IntPtr for aggregation
        public void /* ITypeInfo:: */ GetRefTypeInfo(int hRef, out ComTypes.ITypeInfo ppTI)
        {
            IntPtr typeInfoPtr = IntPtr.Zero;
            ((ITypeInfo_Ptrs)target_ITypeInfo).GetRefTypeInfo(hRef, out typeInfoPtr);
            var outVal = new TypeInfoWrapper(typeInfoPtr, _isUserFormBaseClass ? (int?)hRef : null); // takes ownership of the COM reference
            ppTI = outVal;

            _typeInfosWrapped = _typeInfosWrapped ?? new DisposableList<TypeInfoWrapper>();
            _typeInfosWrapped.Add(outVal);
        }

        public void /* ITypeInfo:: */ GetContainingTypeLib(out ComTypes.ITypeLib ppTLB, out int pIndex)
        {
            ppTLB = _containerTypeLib;
            pIndex = _containerTypeLibIndex;
        }

        // All other members just pass through to the wrappedObject
        public void /* ITypeInfo:: */ GetTypeAttr(out IntPtr ppTypeAttr)
            => target_ITypeInfo.GetTypeAttr(out ppTypeAttr);
        public void /* ITypeInfo:: */ GetTypeComp(out ComTypes.ITypeComp ppTComp)
            => target_ITypeInfo.GetTypeComp(out ppTComp);
        public void /* ITypeInfo:: */ GetFuncDesc(int index, out IntPtr ppFuncDesc)
            => target_ITypeInfo.GetFuncDesc(index, out ppFuncDesc);
        public void /* ITypeInfo:: */ GetVarDesc(int index, out IntPtr ppVarDesc)
            => target_ITypeInfo.GetVarDesc(index, out ppVarDesc);
        public void /* ITypeInfo:: */ GetNames(int memid, string[] rgBstrNames, int cMaxNames, out int pcNames)
            => target_ITypeInfo.GetNames(memid, rgBstrNames, cMaxNames, out pcNames);
        public void /* ITypeInfo:: */ GetRefTypeOfImplType(int index, out int href)
            => target_ITypeInfo.GetRefTypeOfImplType(index, out href);
        public void /* ITypeInfo:: */ GetImplTypeFlags(int index, out ComTypes.IMPLTYPEFLAGS pImplTypeFlags)
            => target_ITypeInfo.GetImplTypeFlags(index, out pImplTypeFlags);
        public void /* ITypeInfo:: */ GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId)
            => target_ITypeInfo.GetIDsOfNames(rgszNames, cNames, pMemId);
        public void /* ITypeInfo:: */ Invoke(object pvInstance, int memid, short wFlags, ref ComTypes.DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr)
            => target_ITypeInfo.Invoke(pvInstance, memid, wFlags, ref pDispParams, pVarResult, pExcepInfo, out puArgErr);
        public void /* ITypeInfo:: */ GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile)
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
        public void /* ITypeInfo:: */ GetDllEntry(int memid, ComTypes.INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
            => target_ITypeInfo.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
        public void /* ITypeInfo:: */ AddressOfMember(int memid, ComTypes.INVOKEKIND invKind, out IntPtr ppv)
            => target_ITypeInfo.AddressOfMember(memid, invKind, out ppv);
        public void /* ITypeInfo:: */ CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj)
            => target_ITypeInfo.CreateInstance(pUnkOuter, riid, out ppvObj);
        public void /* ITypeInfo:: */ GetMops(int memid, out string pBstrMops)
            => target_ITypeInfo.GetMops(memid, out pBstrMops);
        public void /* ITypeInfo:: */ ReleaseTypeAttr(IntPtr pTypeAttr)
            => target_ITypeInfo.ReleaseTypeAttr(pTypeAttr);
        public void /* ITypeInfo:: */ ReleaseFuncDesc(IntPtr pFuncDesc)
            => target_ITypeInfo.ReleaseFuncDesc(pFuncDesc);
        public void /* ITypeInfo:: */ ReleaseVarDesc(IntPtr pVarDesc)
            => target_ITypeInfo.ReleaseVarDesc(pVarDesc);

        public IDispatch GetStdModInstance()
        {
            if (HasVBEExtensions())
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
            if (HasVBEExtensions())
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

        public TypeInfoWrapper GetImplementedTypeInfoByIndex(int implIndex)
        {
            ComTypes.ITypeInfo typeInfoImpl = null;
            int href = 0;
            GetRefTypeOfImplType(implIndex, out href);
            GetRefTypeInfo(href, out typeInfoImpl);
            return (TypeInfoWrapper)typeInfoImpl;
        }

        public bool DoesImplement(string containerName, string interfaceName)
        {
            // check we are not runtime generated with no container
            if (HasNoContainer()) return false;

            if ((containerName == _containerTypeLib.Name) && (Name == interfaceName)) return true;

            for (int implIndex = 0; implIndex < _cachedAttributes.cImplTypes; implIndex++)
            {
                using (var typeInfoImplEx = GetImplementedTypeInfoByIndex(implIndex))
                {
                    if (typeInfoImplEx.DoesImplement(containerName, interfaceName))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public bool DoesImplement(Guid interfaceIID)
        {
            if (_cachedAttributes.guid == interfaceIID) return true;

            for (int implIndex = 0; implIndex < _cachedAttributes.cImplTypes; implIndex++)
            {
                using (var typeInfoImplEx = GetImplementedTypeInfoByIndex(implIndex))
                {
                    if (typeInfoImplEx.DoesImplement(interfaceIID))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public TypeInfoWrapper GetImplementedTypeInfo(string searchTypeName)
        {
            for (int implIndex = 0; implIndex < _cachedAttributes.cImplTypes; implIndex++)
            {
                var typeInfoImplEx = GetImplementedTypeInfoByIndex(implIndex);
                if (typeInfoImplEx.Name == searchTypeName)
                {
                    return typeInfoImplEx;
                }
                typeInfoImplEx.Dispose();
            }

            throw new ArgumentException($"TypeLibWrapper::GetImplementedTypeInfo failed. '{searchTypeName}' module not found.");
        }

        // FIXME this needs work
        // Gets the control ITypeInfo by looking for the corresponding getter on the form interface and returning its retval type
        // Supports UserForms.  what about Access forms etc
        public TypeInfoWrapper GetControlType(string controlName)
        {
            for (int funcIndex = 0; funcIndex < _cachedAttributes.cFuncs; funcIndex++)
            {
                IntPtr funcDescPtr = IntPtr.Zero;
                GetFuncDesc(funcIndex, out funcDescPtr);
                var funcDesc = StructHelper.ReadStructure<ComTypes.FUNCDESC>(funcDescPtr);

                try
                {
                    var names = new string[1];
                    int cNames = 0;
                    GetNames(funcDesc.memid, names, names.Length, out cNames);

                    if ((names[0] == controlName) &&
                            ((funcDesc.invkind & ComTypes.INVOKEKIND.INVOKE_PROPERTYGET) != 0) &&
                            (funcDesc.cParams == 0) &&
                            (funcDesc.elemdescFunc.tdesc.vt == (short)VarEnum.VT_PTR))
                    {
                        var retValElement = StructHelper.ReadStructure<ComTypes.ELEMDESC>(funcDesc.elemdescFunc.tdesc.lpValue);
                        if (retValElement.tdesc.vt == (short)VarEnum.VT_USERDEFINED)
                        {
                            ComTypes.ITypeInfo referenceType;
                            GetRefTypeInfo((int)retValElement.tdesc.lpValue, out referenceType);
                            return (TypeInfoWrapper)referenceType;
                        }                        
                    }
                }
                catch (Exception)
                {
                    // it's fine if GetNames() or GetRefTypeInfo() throws here, we just ignore and move on.
                }                   
                finally
                {
                    ReleaseFuncDesc(funcDescPtr);
                }
            }

            throw new ArgumentException($"TypeInfoWrapper::GetControlType failed. '{controlName}' control not found.");
        }
    }

    // A wrapper for ITypeLib that exposes VBE ITypeInfos safely for managed consumption, plus adds ConditionalCompilationArguments property
    public class TypeLibWrapper : ComTypes.ITypeLib, IDisposable
    {
        private DisposableList<TypeInfoWrapper> _typeInfosWrapped;
        private readonly bool _wrappedObjectIsWeakReference;
        
        private string _name;
        private string _docString;
        private int _helpContext;
        private string _helpFile;

        public string Name { get => _name; }
        public string DocString { get => _docString; }
        public int HelpContext { get => _helpContext; }
        public string HelpFile { get => _helpFile; }

        private ComTypes.ITypeLib target_ITypeLib;
        private IVBEProject target_IVBEProject;

        public bool HasVBEExtensions() => target_IVBEProject != null;

        /*
         This is not yet used, but here in case we want to use this interface at some point.
        private RestrictComInterfaceByAggregation<IVBEProject2> _cachedIVBEProject2;   NEEDS DISPOSING
        private IVBEProjectEx2 target_IVBEProject2
        {
            get
            {
                if (_cachedIVBProjectEx2 == null)
                {
                    if (HasVBEExtensions())
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
            target_IVBEProject = target_ITypeLib as IVBEProject;
            target_ITypeLib.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out _name, out _docString, out _helpContext, out _helpFile);
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

        // We have to wrap the ITypeInfo returned by GetTypeInfo
        // so we cast to our IVBETypeLib interface in order to work with the raw IntPtr for aggregation
        public void /* ITypeLib:: */ GetTypeInfo(int index, out ComTypes.ITypeInfo ppTI)
        {
            IntPtr typeInfoPtr = IntPtr.Zero;
            ((ITypeLib_Ptrs)target_ITypeLib).GetTypeInfo(index, out typeInfoPtr);
            var outVal = new TypeInfoWrapper(typeInfoPtr);
            ppTI = outVal;     // takes ownership of the COM reference

            _typeInfosWrapped = _typeInfosWrapped ?? new DisposableList<TypeInfoWrapper>();
            _typeInfosWrapped.Add(outVal);
        }

        // We have to wrap the ITypeInfo returned by GetTypeInfoOfGuid
        // so we cast to our IVBETypeLib interface in order to work with the raw IntPtr for aggregation
        public void /* ITypeLib:: */ GetTypeInfoOfGuid(ref Guid guid, out ComTypes.ITypeInfo ppTInfo)
        {
            IntPtr typeInfoPtr = IntPtr.Zero;
            ((ITypeLib_Ptrs)target_ITypeLib).GetTypeInfoOfGuid(guid, out typeInfoPtr);
            var outVal = new TypeInfoWrapper(typeInfoPtr);  // takes ownership of the COM reference
            ppTInfo = outVal;

            _typeInfosWrapped = _typeInfosWrapped ?? new DisposableList<TypeInfoWrapper>();
            _typeInfosWrapped.Add(outVal);
        }

        // All other members just pass through to the wrappedObject
        public int /* ITypeLib:: */ GetTypeInfoCount()
            => target_ITypeLib.GetTypeInfoCount();
        public void /* ITypeLib:: */ GetTypeInfoType(int index, out ComTypes.TYPEKIND pTKind)
            => target_ITypeLib.GetTypeInfoType(index, out pTKind);
        public void /* ITypeLib:: */ GetLibAttr(out IntPtr ppTLibAttr)
            => target_ITypeLib.GetLibAttr(out ppTLibAttr);
        public void /* ITypeLib:: */ GetTypeComp(out ComTypes.ITypeComp ppTComp)
            => target_ITypeLib.GetTypeComp(out ppTComp);
        public void /* ITypeLib:: */ GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile)
            => target_ITypeLib.GetDocumentation(index, out strName, out strDocString, out dwHelpContext, out strHelpFile);
        public bool /* ITypeLib:: */ IsName(string szNameBuf, int lHashVal)
            => target_ITypeLib.IsName(szNameBuf, lHashVal);

        // FIXME need to wrap the elements of ITypeInfos returned in FindName here.  RD never calls ITypeInfo::FindName() though, so low priority.
        public void /* ITypeLib:: */ FindName(string szNameBuf, int lHashVal, ComTypes.ITypeInfo[] ppTInfo, int[] rgMemId, ref short pcFound)
            => target_ITypeLib.FindName(szNameBuf, lHashVal, ppTInfo, rgMemId, pcFound);
        public void /* ITypeLib:: */ ReleaseTLibAttr(IntPtr pTLibAttr)
            => target_ITypeLib.ReleaseTLibAttr(pTLibAttr);

        public bool CompileProject()
        {
            if (HasVBEExtensions())
            {
                try
                {
                    target_IVBEProject.CompileProject();
                    return true;
                }
                catch (Exception e)
                {
                    if (e.HResult == (int)VBECompilerConsts.E_VBA_COMPILEERROR)
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

        public string ConditionalCompilationArguments
        {
            get
            {
                if (HasVBEExtensions())
                {
                    return target_IVBEProject.get_ConditionalCompilationArgs();
                }
                else
                {
                    throw new ArgumentException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }
            }

            set
            {
                if (HasVBEExtensions())
                {
                    target_IVBEProject.set_ConditionalCompilationArgs(value);
                }
                else
                {
                    throw new ArgumentException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }
            }
        }

        public TypeInfoWrapper FindTypeInfo(string searchTypeName)
        {
            int countOfTypes = GetTypeInfoCount();

            for (int typeIdx = 0; typeIdx < countOfTypes; typeIdx++)
            {
                ComTypes.ITypeInfo typeInfo;
                GetTypeInfo(typeIdx, out typeInfo);

                var typeInfoEx = (TypeInfoWrapper)typeInfo;
                if (typeInfoEx.Name == searchTypeName)
                {
                    return typeInfoEx;
                }

                typeInfoEx.Dispose();
            }

            throw new ArgumentException($"TypeLibWrapper::FindTypeInfo failed. '{searchTypeName}' module not found.");
        }
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

        public TypeLibWrapper FindTypeLib(string searchLibName)
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
    }
}
