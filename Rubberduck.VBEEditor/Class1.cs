using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Reflection = System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Text;

// create some aliases as there are conflicts between these types in InteropServices and InteropServices.ComTypes
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;
using DISPPARAMS = System.Runtime.InteropServices.ComTypes.DISPPARAMS;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using IVBE = Rubberduck.VBEditor.SafeComWrappers.Abstract.IVBE;

// USAGE GUIDE:   see class VBETypeLibsAPI for demonstrations of usage.
//
// The root object for exposure of the type libraries is TypeLibsAccessor_VBE.  It takes the VBE in its construtor.
// The main wrappers (TypeLibWrapper_VBE and TypeInfoWrapper_VBE) can be used as regular ITypeLib and ITypeInfo objects through casting.
//
// THIS IS A WORK IN PROGRESS.  ERROR HANDLING NEEDS WORK, AS DOES A FEW OF THE HELPER ROUTINES.
//
// WARNING: when using TypeLibsAccessor_VBE directly, do not cache it  
//   The type library is LIVE information, so consider it a snapshot at that very moment when you are dealing with it
//   Make sure you call TypeLibsAccessor_VBE.Dispose() as soon as you have done what you need to do with it.
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

namespace Rubberduck.VBEditor.TypeLibsAPI
{
    public class VBETypeLibsAPI
    {
        public static void ExecuteCode(IVBE ide, string projectName, string standardModuleName, string procName)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                typeLibs.FindTypeLib(projectName).FindTypeInfo(standardModuleName)
                    .StdModExecute(procName, Reflection.BindingFlags.InvokeMethod);
            }
        }

        public static string GetProjectConditionalCompilationArgs(IVBE ide, string projectName)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).ConditionalCompilationArguments;
            }
        }

        public static void SetProjectConditionalCompilationArgs(IVBE ide, string projectName, string newConditionalArgs)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                typeLibs.FindTypeLib(projectName).ConditionalCompilationArguments = newConditionalArgs;
            }
        }

        public static bool IsAWorkbook(IVBE ide, string projectName, string className)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(className).DoesImplement("_Workbook");
            }
        }

        public static bool IsAWorksheet(IVBE ide, string projectName, string className)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(className).DoesImplement("_Worksheet");
            }
        }

        public static string GetUserFormControlType(IVBE ide, string projectName, string userFormName, string controlName)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(userFormName)
                        .GetImplementedTypeInfo("FormItf").GetControlType(controlName).Name;
            }
        }

        public static string DocumentAll(IVBE ide)
        {
            var documenter = new TypeLibDocumenter();

            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                foreach (var typeLib in typeLibs)
                {
                    documenter.AddTypeLib(typeLib);
                }
            }

            return documenter.ToString();
        }
    }

    public enum TYPEKINDEx
    {
        TKIND_ENUM = 0,
        TKIND_RECORD = 1,
        TKIND_MODULE = 2,
        TKIND_INTERFACE = 3,
        TKIND_DISPATCH = 4,
        TKIND_COCLASS = 5,
        TKIND_ALIAS = 6,
        TKIND_UNION = 7,

        TKIND_VBACLASS = 8,                 // extended by VBA, this is used for the outermost interface
    }

    public class StructHelper
    {
        public static T ReadStructure<T>(object comObj)
        {
            // Reads a COM object as a structure to copy its internal fields
            var referencesPtr = Marshal.GetIUnknownForObject(comObj);
            var retVal = StructHelper.ReadStructure<T>(referencesPtr);
            Marshal.Release(referencesPtr);
            return retVal;
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

    // An internal representation of the VBE References collection object, as returned from the VBE.ActiveVBProject.References, or similar
    // These offsets are known to be valid across 32-bit and 64-bit versions of VBA and VB6, right back from when VBA6 was first released.
    [StructLayout(LayoutKind.Sequential)]
    struct ReferencesObj_VBE
    {
        IntPtr vTable1;     // _References vtable
        IntPtr vTable2;
        IntPtr vTable3;
        IntPtr Object1;
        IntPtr Object2;
        public IntPtr TypeLib;
        IntPtr Placeholder1;
        IntPtr Placeholder2;
        IntPtr RefCount;
    }

    // A ITypeLib object hosted by the VBE, also providing Prev/Next pointers for a double linked list of all loaded project ITypeLibs
    [StructLayout(LayoutKind.Sequential)]
    struct TypeLibObj_VBE
    {
        IntPtr vTable1;     // ITypeLib vtable
        IntPtr vTable2;
        IntPtr vTable3;
        public IntPtr Prev;
        public IntPtr Next;
    }

    // IVBProjectEx_VBE, obtainable from a VBE hosted ITypeLib in order to access a few extra features...
    [ComImport(), Guid("DDD557E0-D96F-11CD-9570-00AA0051E5D4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IVBProjectEx_VBE
    {
        void Placeholder1();
        void Placeholder2();
        int VBE_LCID();
        void Placeholder3();
        void Placeholder4();
        void Placeholder5();
        void Placeholder6();
        void Placeholder7();
        string get_ConditionalCompilationArgs();
        void set_ConditionalCompilationArgs(string args);
    }

    // AggregateSingleInterface is used to ensure that a wrapped COM object only responds to a specific interface
    // In particular, we don't want them to respond to IProvideClassInfo, which is broken in the VBE for some ITypeInfo implementations 
    public class AggregateSingleInterface<T> : ICustomQueryInterface, IDisposable
        where T : class
    {
        private IntPtr _outerObject;

        public AggregateSingleInterface(IntPtr outerObject)
        {
            _outerObject = outerObject;
            Marshal.AddRef(_outerObject);
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            Marshal.Release(_outerObject);
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            if ((!_isDisposed) && (iid == typeof(T).GUID))       // no need to offer IID_IUnknown here, as it is handled by the aggregation object
            {
                ppv = _outerObject;
                Marshal.AddRef(_outerObject);
                return CustomQueryInterfaceResult.Handled;
            }
            return CustomQueryInterfaceResult.Failed;
        }
    }

    [ComImport(), Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDispatch
    {
    }

    // A compatible version of ITypeInfo, where COM objects are outputted as IntPtrs instead of objects
    [ComImport(), Guid("00020401-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITypeInfo_VBE
    {
        void GetTypeAttr(out IntPtr ppTypeAttr);
        void GetTypeComp(out IntPtr ppTComp);
        void GetFuncDesc(int index, out IntPtr ppFuncDesc);
        void GetVarDesc(int index, out IntPtr ppVarDesc);
        void GetNames(int memid, [Out] out string rgBstrNames, int cMaxNames, out int pcNames);
        void GetRefTypeOfImplType(int index, out int href);
        void GetImplTypeFlags(int index, out IMPLTYPEFLAGS pImplTypeFlags);
        void GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId);
        void Invoke(object pvInstance, int memid, short wFlags, ref DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr);
        void GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile);
        void GetDllEntry(int memid, INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal);
        void GetRefTypeInfo(int hRef, out IntPtr ppTI);
        void AddressOfMember(int memid, INVOKEKIND invKind, out IntPtr ppv);
        void CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj);
        void GetMops(int memid, out string pBstrMops);
        void GetContainingTypeLib(out IntPtr ppTLB, out int pIndex);
        void ReleaseTypeAttr(IntPtr pTypeAttr);
        void ReleaseFuncDesc(IntPtr pFuncDesc);
        void ReleaseVarDesc(IntPtr pVarDesc);

        void Placeholder1();
        IDispatch GetStdModInstance();            // a handy extra vtable entry we can use to invoke members in standard modules.
    }

    // FIXME there's probably some better builtin c# class for this
    public class DisposableList<T> : List<T>, IDisposable
        where T : class
    {
        public void Dispose() => ((IDisposable)this).Dispose();

        private bool _isDisposed;
        void IDisposable.Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            foreach (var element in this)
            {
                ((IDisposable)element).Dispose();
            }
        }
    }

    public enum TypeLibConsts : int
    {
        MEMBERID_NIL = -1,
    }

    // A wrapper for ITypeInfo provided by VBE, allowing safe managed consumption, plus adds StdModExecute functionality
    public class TypeInfoWrapper_VBE : ITypeInfo, IDisposable
    {
        private DisposableList<TypeInfoWrapper_VBE> _typeInfosWrapped;
        private DisposableList<TypeLibWrapper_VBE> _typeLibsWrapped;
        private AggregateSingleInterface<ITypeInfo> _typeInfoAggregatorObj;
        private bool _isUserFormBaseClass = false;
        private int _firstImplementedTypeHref = -1;
        private TYPEATTR _cachedAttributes;

        public readonly string Name;
        public readonly string DocString;
        public readonly int HelpContext;
        public readonly string HelpFile;

        private ITypeInfo _wrappedObject;
        private ITypeInfo_VBE _ITypeInfoAlt
        { get => ((ITypeInfo_VBE)_wrappedObject); }

        public TypeInfoWrapper_VBE(IntPtr rawObjectPtr, bool isBrokenUserFormBaseClass = false, bool isBrokenUserFormBaseEventsClass = false)
        {
            _typeInfoAggregatorObj = new AggregateSingleInterface<ITypeInfo>(rawObjectPtr);
            var aggObjPtr = Marshal.CreateAggregatedObject(rawObjectPtr, _typeInfoAggregatorObj);
            _wrappedObject = (ITypeInfo)Marshal.GetObjectForIUnknown(aggObjPtr);        // when this CCW object gets released, it will free the aggObjInner (well, after GC)
            Marshal.Release(aggObjPtr);         // _wrappedObject holds a reference to this now

            IntPtr typeAttrPtr = IntPtr.Zero;
            GetTypeAttr(out typeAttrPtr);
            _cachedAttributes = StructHelper.ReadStructure<TYPEATTR>(typeAttrPtr);
            ReleaseTypeAttr(typeAttrPtr);      // don't need to keep a hold of it, as _cachedAttributes is a copy

            if (isBrokenUserFormBaseClass)
            {
                Name = "_UserFormBase{unnamed}";
            }
            else if (isBrokenUserFormBaseEventsClass)
            {
                Name = "_UserFormBaseEvents{unnamed}";
            }
            else
            {
                _wrappedObject.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out Name, out DocString, out HelpContext, out HelpFile);

                // Determine if this is a UserForm base class, that requires special handling to workaround a VBE bug in its implemented classes
                // the guids are dynamic, so we can't use them for detection.
                if ((_cachedAttributes.typekind == TYPEKIND.TKIND_COCLASS) && (Name == "Form") && (_cachedAttributes.cImplTypes == 2))
                {
                    _isUserFormBaseClass = true;
                }
            }
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            if (_typeInfosWrapped != null) _typeInfosWrapped.Dispose();
            if (_typeLibsWrapped != null) _typeLibsWrapped.Dispose();
            Marshal.ReleaseComObject(_wrappedObject);
            _typeInfoAggregatorObj.Dispose();
        }

        // We have to wrap the ITypeInfo returned by GetRefTypeInfo
        // so we cast to our ITypeInfo_VBE interface in order to work with the raw IntPtr for aggregation
        public void /* ITypeInfo:: */ GetRefTypeInfo(int hRef, out ITypeInfo ppTI)
        {
            IntPtr typeInfoPtr = IntPtr.Zero;
            _ITypeInfoAlt.GetRefTypeInfo(hRef, out typeInfoPtr);
            var isBrokenUserFormBaseClass = _isUserFormBaseClass && (hRef == _firstImplementedTypeHref);
            var isBrokenUserFormBaseEventsClass = _isUserFormBaseClass && (hRef != _firstImplementedTypeHref);
            var outVal = new TypeInfoWrapper_VBE(typeInfoPtr, isBrokenUserFormBaseClass, isBrokenUserFormBaseEventsClass); // takes ownership of the COM reference
            ppTI = outVal;

            if (_typeInfosWrapped == null) _typeInfosWrapped = new DisposableList<TypeInfoWrapper_VBE>();
            _typeInfosWrapped.Add(outVal);
        }

        // We have to wrap the ITypeLib returned by GetContainingTypeLib
        // so we cast to our ITypeInfo_VBE interface in order to work with the raw IntPtr for aggregation
        public void /* ITypeInfo:: */ GetContainingTypeLib(out ITypeLib ppTLB, out int pIndex)
        {
            IntPtr typeLibPtr = IntPtr.Zero;
            _ITypeInfoAlt.GetContainingTypeLib(out typeLibPtr, out pIndex);
            var outVal = new TypeLibWrapper_VBE(typeLibPtr);  // takes ownership of the COM reference
            ppTLB = outVal;

            if (_typeLibsWrapped == null) _typeLibsWrapped = new DisposableList<TypeLibWrapper_VBE>();
            _typeLibsWrapped.Add(outVal);
        }

        // All other members just pass through to the wrappedObject
        public void /* ITypeInfo:: */ GetTypeAttr(out IntPtr ppTypeAttr)
            => _wrappedObject.GetTypeAttr(out ppTypeAttr);
        public void /* ITypeInfo:: */ GetTypeComp(out ITypeComp ppTComp)
            => _wrappedObject.GetTypeComp(out ppTComp);
        public void /* ITypeInfo:: */ GetFuncDesc(int index, out IntPtr ppFuncDesc)
            => _wrappedObject.GetFuncDesc(index, out ppFuncDesc);
        public void /* ITypeInfo:: */ GetVarDesc(int index, out IntPtr ppVarDesc)
            => _wrappedObject.GetVarDesc(index, out ppVarDesc);
        public void /* ITypeInfo:: */ GetNames(int memid, string[] rgBstrNames, int cMaxNames, out int pcNames)
            => _wrappedObject.GetNames(memid, rgBstrNames, cMaxNames, out pcNames);
        public void /* ITypeInfo:: */ GetRefTypeOfImplType(int index, out int href)
        {
            _wrappedObject.GetRefTypeOfImplType(index, out href);
            if (index == 0) _firstImplementedTypeHref = href;
        }
        public void /* ITypeInfo:: */ GetImplTypeFlags(int index, out IMPLTYPEFLAGS pImplTypeFlags)
            => _wrappedObject.GetImplTypeFlags(index, out pImplTypeFlags);
        public void /* ITypeInfo:: */ GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId)
            => _wrappedObject.GetIDsOfNames(rgszNames, cNames, pMemId);
        public void /* ITypeInfo:: */ Invoke(object pvInstance, int memid, short wFlags, ref DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr)
            => _wrappedObject.Invoke(pvInstance, memid, wFlags, ref pDispParams, pVarResult, pExcepInfo, out puArgErr);
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
                _wrappedObject.GetDocumentation(index, out strName, out strDocString, out dwHelpContext, out strHelpFile);
            }
        }
        public void /* ITypeInfo:: */ GetDllEntry(int memid, INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
            => _wrappedObject.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);
        public void /* ITypeInfo:: */ AddressOfMember(int memid, INVOKEKIND invKind, out IntPtr ppv)
            => _wrappedObject.AddressOfMember(memid, invKind, out ppv);
        public void /* ITypeInfo:: */ CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj)
            => _wrappedObject.CreateInstance(pUnkOuter, riid, out ppvObj);
        public void /* ITypeInfo:: */ GetMops(int memid, out string pBstrMops)
            => _wrappedObject.GetMops(memid, out pBstrMops);
        public void /* ITypeInfo:: */ ReleaseTypeAttr(IntPtr pTypeAttr)
            => _wrappedObject.ReleaseTypeAttr(pTypeAttr);
        public void /* ITypeInfo:: */ ReleaseFuncDesc(IntPtr pFuncDesc)
            => _wrappedObject.ReleaseFuncDesc(pFuncDesc);
        public void /* ITypeInfo:: */ ReleaseVarDesc(IntPtr pVarDesc)
            => _wrappedObject.ReleaseVarDesc(pVarDesc);

        public IDispatch GetStdModInstance() => _ITypeInfoAlt.GetStdModInstance();
        public object StdModExecute(string name, Reflection.BindingFlags invokeAttr, object[] args = null)
        {
            var StaticModule = GetStdModInstance();
            var retVal = StaticModule.GetType().InvokeMember(name, invokeAttr, null, StaticModule, args);
            Marshal.ReleaseComObject(StaticModule);
            return retVal;
        }

        public TypeInfoWrapper_VBE GetImplementedTypeInfoByIndex(int implIndex)
        {
            ITypeInfo typeInfoImpl = null;
            int href = 0;
            GetRefTypeOfImplType(implIndex, out href);
            GetRefTypeInfo(href, out typeInfoImpl);
            return (TypeInfoWrapper_VBE)typeInfoImpl;
        }

        public bool DoesImplement(string interfaceName)
        {
            if (Name == interfaceName) return true;

            for (int implIndex = 0; implIndex < _cachedAttributes.cImplTypes; implIndex++)
            {
                using (var typeInfoImplEx = GetImplementedTypeInfoByIndex(implIndex))
                {
                    if (typeInfoImplEx.DoesImplement(interfaceName))
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

        public TypeInfoWrapper_VBE GetImplementedTypeInfo(string searchTypeName)
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

            throw new ArgumentException($"TypeLibWrapper_VBE::GetImplementedTypeInfo failed. '{searchTypeName}' module not found.");
        }

        // FIXME this needs work
        // Gets the control ITypeInfo by looking for the corresponding getter on the form interface and returning its retval type
        // Supports UserForms.  what about Access forms etc
        public TypeInfoWrapper_VBE GetControlType(string controlName)
        {
            for (int funcIndex = 0; funcIndex < _cachedAttributes.cFuncs; funcIndex++)
            {
                IntPtr funcDescPtr = IntPtr.Zero;
                GetFuncDesc(funcIndex, out funcDescPtr);
                var funcDesc = StructHelper.ReadStructure<FUNCDESC>(funcDescPtr);

                try
                {
                    var names = new string[1];
                    int cNames = 0;
                    GetNames(funcDesc.memid, names, names.Length, out cNames);

                    if (names[0] == controlName)
                    {
                        if (((funcDesc.invkind & INVOKEKIND.INVOKE_PROPERTYGET) != 0) && (funcDesc.cParams == 0))
                        {
                            if (funcDesc.elemdescFunc.tdesc.vt == 26)       // VT_PTR
                            {
                                var retValElement = StructHelper.ReadStructure<ELEMDESC>(funcDesc.elemdescFunc.tdesc.lpValue);
                                if (retValElement.tdesc.vt == 29)       // VT_USERDEFINED
                                {
                                    ITypeInfo referenceType;
                                    GetRefTypeInfo((int)retValElement.tdesc.lpValue, out referenceType);
                                    return (TypeInfoWrapper_VBE)referenceType;
                                }
                            }
                        }
                    }
                }
                finally
                {
                    ReleaseFuncDesc(funcDescPtr);
                }
            }

            throw new ArgumentException($"TypeInfoWrapper_VBE::GetControlType failed. '{controlName}' control not found.");
        }
    }

    // A compatible version of ITypeLib, where COM objects are outputted as IntPtrs instead of objects
    [ComImport(), Guid("00020402-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITypeLib_VBE
    {
        int GetTypeInfoCount();
        void GetTypeInfo(int index, out IntPtr ppTI);
        void GetTypeInfoType(int index, out TYPEKIND pTKind);
        void GetTypeInfoOfGuid(ref Guid guid, out IntPtr ppTInfo);
        void GetLibAttr(out IntPtr ppTLibAttr);
        void GetTypeComp(out IntPtr ppTComp);
        void GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile);
        bool IsName(string szNameBuf, int lHashVal);
        void FindName(string szNameBuf, int lHashVal, IntPtr[] ppTInfo, int[] rgMemId, ref short pcFound);
        void ReleaseTLibAttr(IntPtr pTLibAttr);
    }

    // A wrapper for ITypeLib that exposes VBE ITypeInfos safely for managed consumption, plus adds ConditionalCompilationArguments property
    public class TypeLibWrapper_VBE : ITypeLib, IDisposable
    {
        private DisposableList<TypeInfoWrapper_VBE> _typeInfosWrapped;
        private ITypeLib _wrappedObject;

        public readonly string Name;
        public readonly string DocString;
        public readonly int HelpContext;
        public readonly string HelpFile;

        private ITypeLib_VBE _ITypeLibAlt
        { get => ((ITypeLib_VBE)_wrappedObject); }

        private IVBProjectEx_VBE _IVBProjectEx
        { get => ((IVBProjectEx_VBE)_wrappedObject); }

        public TypeLibWrapper_VBE(IntPtr rawObjectPtr)
        {
            _wrappedObject = (ITypeLib)Marshal.GetObjectForIUnknown(rawObjectPtr);
            Marshal.Release(rawObjectPtr);         // _wrappedObject holds a reference to this now

            _wrappedObject.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out Name, out DocString, out HelpContext, out HelpFile);
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            if (_typeInfosWrapped != null) _typeInfosWrapped.Dispose();
            Marshal.ReleaseComObject(_wrappedObject);
        }

        // We have to wrap the ITypeInfo returned by GetTypeInfo
        // so we cast to our ITypeLib_VBE interface in order to work with the raw IntPtr for aggregation
        public void /* ITypeLib:: */ GetTypeInfo(int index, out ITypeInfo ppTI)
        {
            IntPtr typeInfoPtr = IntPtr.Zero;
            _ITypeLibAlt.GetTypeInfo(index, out typeInfoPtr);
            var outVal = new TypeInfoWrapper_VBE(typeInfoPtr);
            ppTI = outVal;     // takes ownership of the COM reference

            if (_typeInfosWrapped == null) _typeInfosWrapped = new DisposableList<TypeInfoWrapper_VBE>();
            _typeInfosWrapped.Add(outVal);
        }

        // We have to wrap the ITypeInfo returned by GetTypeInfoOfGuid
        // so we cast to our ITypeLib_VBE interface in order to work with the raw IntPtr for aggregation
        public void /* ITypeLib:: */ GetTypeInfoOfGuid(ref Guid guid, out ITypeInfo ppTInfo)
        {
            IntPtr typeInfoPtr = IntPtr.Zero;
            _ITypeLibAlt.GetTypeInfoOfGuid(guid, out typeInfoPtr);
            var outVal = new TypeInfoWrapper_VBE(typeInfoPtr);  // takes ownership of the COM reference
            ppTInfo = outVal;

            if (_typeInfosWrapped == null) _typeInfosWrapped = new DisposableList<TypeInfoWrapper_VBE>();
            _typeInfosWrapped.Add(outVal);
        }

        // All other members just pass through to the wrappedObject
        public int /* ITypeLib:: */ GetTypeInfoCount()
            => _wrappedObject.GetTypeInfoCount();
        public void /* ITypeLib:: */ GetTypeInfoType(int index, out TYPEKIND pTKind)
            => _wrappedObject.GetTypeInfoType(index, out pTKind);
        public void /* ITypeLib:: */ GetLibAttr(out IntPtr ppTLibAttr)
            => _wrappedObject.GetLibAttr(out ppTLibAttr);
        public void /* ITypeLib:: */ GetTypeComp(out ITypeComp ppTComp)
            => _wrappedObject.GetTypeComp(out ppTComp);
        public void /* ITypeLib:: */ GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile)
            => _wrappedObject.GetDocumentation(index, out strName, out strDocString, out dwHelpContext, out strHelpFile);
        public bool /* ITypeLib:: */ IsName(string szNameBuf, int lHashVal)
            => _wrappedObject.IsName(szNameBuf, lHashVal);

        // FIXME need to wrap the elements of ITypeInfos returned in FindName here.  RD never calls ITypeInfo::FindName() though, so low priority.
        public void /* ITypeLib:: */ FindName(string szNameBuf, int lHashVal, ITypeInfo[] ppTInfo, int[] rgMemId, ref short pcFound)
            => _wrappedObject.FindName(szNameBuf, lHashVal, ppTInfo, rgMemId, pcFound);
        public void /* ITypeLib:: */ ReleaseTLibAttr(IntPtr pTLibAttr)
            => _wrappedObject.ReleaseTLibAttr(pTLibAttr);

        public string ConditionalCompilationArguments
        {
            get => _IVBProjectEx.get_ConditionalCompilationArgs();
            set => _IVBProjectEx.set_ConditionalCompilationArgs(value);
        }

        public TypeInfoWrapper_VBE FindTypeInfo(string searchTypeName)
        {
            int countOfTypes = GetTypeInfoCount();

            for (int typeIdx = 0; typeIdx < countOfTypes; typeIdx++)
            {
                ITypeInfo typeInfo;
                GetTypeInfo(typeIdx, out typeInfo);

                var typeInfoEx = (TypeInfoWrapper_VBE)typeInfo;
                if (typeInfoEx.Name == searchTypeName)
                {
                    return typeInfoEx;
                }

                typeInfoEx.Dispose();
            }

            throw new ArgumentException($"TypeLibWrapper_VBE::FindTypeInfo failed. '{searchTypeName}' module not found.");
        }
    }

    // class for iterating over the double linked list of ITypeLibs provided by the VBE
    public class TypeLibsIterator_VBE : IEnumerable<TypeLibWrapper_VBE>, IEnumerator<TypeLibWrapper_VBE>, IDisposable
    {
        private IntPtr _currentTypeLibPtr;
        private TypeLibObj_VBE _currentTypeLibStruct;
        private bool _isStart;

        public TypeLibsIterator_VBE(IntPtr typeLibPtr)
        {
            _currentTypeLibPtr = typeLibPtr;
            _currentTypeLibStruct = StructHelper.ReadStructureSafe<TypeLibObj_VBE>(_currentTypeLibPtr);
            Reset();
        }

        public void Dispose()
        {
            // nothing to do here, we don't own anything that needs releasing
        }

        IEnumerator IEnumerable.GetEnumerator() => this;
        public IEnumerator<TypeLibWrapper_VBE> GetEnumerator() => this;

        public IntPtr GetCurrentReference()
        {
            Marshal.AddRef(_currentTypeLibPtr);
            return _currentTypeLibPtr;
        }

        TypeLibWrapper_VBE IEnumerator<TypeLibWrapper_VBE>.Current => new TypeLibWrapper_VBE(GetCurrentReference());
        object IEnumerator.Current => new TypeLibWrapper_VBE(GetCurrentReference());

        public void Reset()  // walk back to the first project in the chain
        {
            while (_currentTypeLibStruct.Prev != IntPtr.Zero)
            {
                _currentTypeLibPtr = _currentTypeLibStruct.Prev;
                _currentTypeLibStruct = StructHelper.ReadStructureSafe<TypeLibObj_VBE>(_currentTypeLibPtr);
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
            _currentTypeLibStruct = StructHelper.ReadStructureSafe<TypeLibObj_VBE>(_currentTypeLibPtr);
            return true;
        }
    }

    // the main class for hooking into the live ITypeLibs provided by the VBE
    public class TypeLibsAccessor_VBE : DisposableList<TypeLibWrapper_VBE>, IDisposable
    {
        public TypeLibsAccessor_VBE(IVBE ide)
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
                            var internalReferencesObj = StructHelper.ReadStructure<ReferencesObj_VBE>(references.Target);

                            // Now we've got this one internalReferencesObj.typeLib, we can iterate through ALL loaded project TypeLibs
                            using (var typeLibIterator = new TypeLibsIterator_VBE(internalReferencesObj.TypeLib))
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
                    finally
                    {
                        // probably a protected project, just move on to the next project.
                    }
                }
            }

            // return an empty list on error
        }

        public TypeLibWrapper_VBE FindTypeLib(string searchLibName)
        {
            foreach (var typeLib in this)
            {
                if (typeLib.Name == searchLibName)
                {
                    return typeLib;
                }
            }

            throw new ArgumentException($"TypeLibsAccessor_VBE::FindTypeLib failed. '{searchLibName}' project not found.");
        }
    }

    // for debug purposes, just reinventing the wheel here to document the major things exposed by a particular ITypeLib
    // (compatible with all ITypeLibs, not just VBE ones, but also documents the VBE specific extensions)
    // this is a throw away class, once proper integration into RD has been achieved.
    public class TypeLibDocumenter
    {
        StringBuilder _document = new StringBuilder();

        public override string ToString() => _document.ToString();

        private void AppendLine(string value = "")
            => _document.Append(value + "\r\n");

        private void AppendLineButRemoveEmbeddedNullChars(string value)
            => AppendLine(value.Replace("\0", string.Empty));

        public void AddTypeLib(ITypeLib typeLib)
        {
            AppendLine();
            AppendLine("================================================================================");
            AppendLine();

            string libName;
            string libString;
            int libHelp;
            string libHelpFile;
            typeLib.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out libName, out libString, out libHelp, out libHelpFile);

            if (libName == null) libName = "[VBA.Immediate.Window]";

            AppendLine("ITypeLib: " + libName);
            if (libString != null) AppendLineButRemoveEmbeddedNullChars("- Documentation: " + libString);
            if (libHelp != 0) AppendLineButRemoveEmbeddedNullChars("- HelpContext: " + libHelp);
            if (libHelpFile != null) AppendLineButRemoveEmbeddedNullChars("- HelpFile: " + libHelpFile);

            IntPtr typeLibAttributesPtr;
            typeLib.GetLibAttr(out typeLibAttributesPtr);
            var typeLibAttributes = StructHelper.ReadStructure<TYPELIBATTR>(typeLibAttributesPtr);
            typeLib.ReleaseTLibAttr(typeLibAttributesPtr);          // no need to keep open.  copied above

            AppendLine("- Guid: " + typeLibAttributes.guid);
            AppendLine("- Lcid: " + typeLibAttributes.lcid);
            AppendLine("- SysKind: " + typeLibAttributes.syskind);
            AppendLine("- LibFlags: " + typeLibAttributes.wLibFlags);
            AppendLine("- MajorVer: " + typeLibAttributes.wMajorVerNum);
            AppendLine("- MinorVer: " + typeLibAttributes.wMinorVerNum);

            var typeLibVBE = typeLib as TypeLibWrapper_VBE;
            if (typeLibVBE != null)
            {
                // This is a VBE ITypeInfo, so we should be able to obtain the conditional compilation arguments
                AppendLine("- VBE Conditional Compilation Arguments: " + typeLibVBE.ConditionalCompilationArguments);
            }

            int CountOfTypes = typeLib.GetTypeInfoCount();
            AppendLine("- TypeCount: " + CountOfTypes);

            for (int typeIdx = 0; typeIdx < CountOfTypes; typeIdx++)
            {
                ITypeInfo typeInfo;
                typeLib.GetTypeInfo(typeIdx, out typeInfo);

                AddTypeInfo(typeInfo, libName, 0);
            }
        }

        void AddTypeInfo(ITypeInfo typeInfo, string qualifiedName, int implementsLevel)
        {
            AppendLine();
            if (implementsLevel == 0)
            {
                AppendLine("-------------------------------------------------------------------------------");
                AppendLine();
            }
            implementsLevel++;

            IntPtr typeAttrPtr = IntPtr.Zero;
            typeInfo.GetTypeAttr(out typeAttrPtr);
            var typeInfoAttributes = StructHelper.ReadStructure<TYPEATTR>(typeAttrPtr);
            typeInfo.ReleaseTypeAttr(typeAttrPtr);

            string typeName = null;
            string typeString = null;
            int typeHelp = 0;
            string TypeHelpFile = null;

            typeInfo.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out typeName, out typeString, out typeHelp, out TypeHelpFile);

            AppendLine(qualifiedName + "::" + (typeName.Replace("\0", string.Empty) ?? "[unnamed]"));
            if (typeString != null) AppendLineButRemoveEmbeddedNullChars("- Documentation: " + typeString.Replace("\0", string.Empty));
            if (typeHelp != 0) AppendLineButRemoveEmbeddedNullChars("- HelpContext: " + typeHelp);
            if (TypeHelpFile != null) AppendLineButRemoveEmbeddedNullChars("- HelpFile: " + TypeHelpFile.Replace("\0", string.Empty));

            AppendLine("- Type: " + (TYPEKINDEx)typeInfoAttributes.typekind);
            AppendLine("- Guid: {" + typeInfoAttributes.guid + "}");

            AppendLine("- cImplTypes (implemented interfaces count): " + typeInfoAttributes.cImplTypes);
            AppendLine("- cFuncs (function count): " + typeInfoAttributes.cFuncs);
            AppendLine("- cVars (fields count): " + typeInfoAttributes.cVars);

            for (int funcIdx = 0; funcIdx < typeInfoAttributes.cFuncs; funcIdx++)
            {
                AddFunc(typeInfo, funcIdx);
            }

            for (int varIdx = 0; varIdx < typeInfoAttributes.cVars; varIdx++)
            {
                AddField(typeInfo, varIdx);
            }

            for (int implIndex = 0; implIndex < typeInfoAttributes.cImplTypes; implIndex++)
            {
                ITypeInfo typeInfoImpl = null;
                int href = 0;
                typeInfo.GetRefTypeOfImplType(implIndex, out href);
                typeInfo.GetRefTypeInfo(href, out typeInfoImpl);

                AppendLine("implements...");
                AddTypeInfo(typeInfoImpl, qualifiedName + "::" + typeName, implementsLevel);
            }
        }

        void AddFunc(ITypeInfo typeInfo, int funcIndex)
        {
            IntPtr funcDescPtr = IntPtr.Zero;
            typeInfo.GetFuncDesc(funcIndex, out funcDescPtr);
            var funcDesc = StructHelper.ReadStructure<FUNCDESC>(funcDescPtr);

            var names = new string[255];
            int cNames = 0;
            typeInfo.GetNames(funcDesc.memid, names, names.Length, out cNames);

            string namesInfo = names[0] + "(";

            int argIndex = 1;
            while (argIndex < cNames)
            {
                if (argIndex > 1) namesInfo += ", ";
                namesInfo += names[argIndex].Length > 0 ? names[argIndex] : "retVal";
                argIndex++;
            }

            namesInfo += ")";

            typeInfo.ReleaseFuncDesc(funcDescPtr);

            AppendLine("- member: " + namesInfo + " [id 0x" + funcDesc.memid.ToString("X") + ", " + funcDesc.invkind + "]");
        }

        void AddField(ITypeInfo typeInfo, int varIndex)
        {
            IntPtr varDescPtr = IntPtr.Zero;
            typeInfo.GetVarDesc(varIndex, out varDescPtr);
            var varDesc = StructHelper.ReadStructure<VARDESC>(varDescPtr);

            if (varDesc.memid != (int)TypeLibConsts.MEMBERID_NIL)
            {
                var names = new string[1];
                int cNames = 0;
                typeInfo.GetNames(varDesc.memid, names, names.Length, out cNames);
                AppendLine("- field: " + names[0] + " [id 0x" + varDesc.memid.ToString("X") + "]");
            }
            else
            {
                // Constants appear in the typelib with no name
                AppendLine("- constant: {unknown name}");
            }

            typeInfo.ReleaseVarDesc(varDescPtr);
        }
    }
}