using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibsSupport;
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
    public sealed class TypeLibWrapper : ITypeLibWrapper
    {
        private DisposableList<TypeInfoWrapper> _typeInfosWrapped;
        private readonly bool _wrappedObjectIsWeakReference;
        public TypeInfosCollection TypeInfos { get; private set; }
        
        /// <summary>
        /// Exposes an enumerable collection of references used by the VBE type library
        /// </summary>
        public class ReferencesCollection : IIndexedCollectionBase<TypeInfoReference>
        {
            private readonly TypeLibWrapper _parent;
            public ReferencesCollection(TypeLibWrapper parent) => _parent = parent;
            public override int Count => _parent.GetVBEReferencesCount();
            public override TypeInfoReference GetItemByIndex(int index) => _parent.GetVBEReferenceByIndex(index);
        }
        public ReferencesCollection VBEReferences;

        private TypeLibTextFields? _cachedTextFields;

        private TypeLibTextFields CachedTextFields
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

        public string Name => CachedTextFields._name;
        public string DocString => CachedTextFields._docString;
        public int HelpContext => CachedTextFields._helpContext;
        public string HelpFile => CachedTextFields._helpFile;

        private readonly ComTypes.ITypeLib target_ITypeLib;
        private IVBEProject target_IVBEProject;

        public bool HasVBEExtensions => target_IVBEProject != null;

        public int GetVBEReferencesCount()
        {
            if (!HasVBEExtensions)
            {
                throw new ArgumentException("This TypeLib does not represent a VBE project, so we cannot get reference strings from it");
            }
            return target_IVBEProject.GetReferencesCount();
        }

        public TypeInfoReference GetVBEReferenceByIndex(int index)
        {
            if (!HasVBEExtensions)
            {
                throw new ArgumentException("This TypeLib does not represent a VBE project, so we cannot get reference strings from it");
            }

            if (index >= target_IVBEProject.GetReferencesCount())
            {
                throw new ArgumentException($"Specified index not valid for the references collection (reference {index} in project {Name})");
            }

            return new TypeInfoReference(this, index, target_IVBEProject.GetReferenceString(index));
        }

        public TypeLibWrapper GetVBEReferenceTypeLibByIndex(int index)
        {
            if (!HasVBEExtensions)
            {
                throw new ArgumentException("This TypeLib does not represent a VBE project, so we cannot get reference strings from it");
            }

            if (index >= target_IVBEProject.GetReferencesCount())
            {
                throw new ArgumentException($"Specified index not valid for the references collection (reference {index} in project {Name})");
            }

            IntPtr referenceTypeLibPtr = target_IVBEProject.GetReferenceTypeLib(index);
            if (referenceTypeLibPtr == IntPtr.Zero)
            {
                throw new ArgumentException("Reference TypeLib not available - probably a missing reference.");
            }
            return new TypeLibWrapper(referenceTypeLibPtr, isRefCountedInput: true);
        }

        public TypeInfoReference GetVBEReferenceByGuid(Guid referenceGuid)
        {
            if (!HasVBEExtensions)
            {
                throw new ArgumentException("This TypeLib does not represent a VBE project, so we cannot get reference strings from it");
            }

            foreach (var reference in VBEReferences)
            {
                if (reference.GUID == referenceGuid)
                {
                    return reference;
                }
            }

            throw new ArgumentException($"Specified GUID not found in references collection {referenceGuid}.");
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

        
        internal static TypeLibWrapper FromVBProject(IVBProject vbProject)
        {
            using (var references = vbProject.References)
            {
                // Now we've got the references object, we can read the internal object structure to grab the ITypeLib
                var internalReferencesObj = StructHelper.ReadComObjectStructure<VBEReferencesObj>(references.Target);
                return new TypeLibWrapper(internalReferencesObj.TypeLib, isRefCountedInput: false);
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
        public TypeLibWrapper(IntPtr rawObjectPtr, bool isRefCountedInput)
        {
            if (!UnmanagedMemHelper.ValidateComObject(rawObjectPtr))
            {
                throw new ArgumentException("Expected COM object, but validation failed.");
            };
            target_ITypeLib = (ComTypes.ITypeLib)Marshal.GetObjectForIUnknown(rawObjectPtr);
            if (isRefCountedInput) Marshal.Release(rawObjectPtr);         
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
            // We cast to our IVBETypeLib interface in order to work with the raw IntPtr for aggregation
            ((ITypeLib_Ptrs)target_ITypeLib).GetTypeInfo(index, out var typeInfoPtr);
            var outVal = new TypeInfoWrapper(typeInfoPtr);
            _typeInfosWrapped?.Dispose();
            _typeInfosWrapped = _typeInfosWrapped ?? new DisposableList<TypeInfoWrapper>();
            _typeInfosWrapped.Add(outVal);
            return outVal;
        }

        public int TypesCount => target_ITypeLib.GetTypeInfoCount();

        private ComTypes.TYPELIBATTR? _cachedLibAttribs;
        public ComTypes.TYPELIBATTR Attributes
        {
            get
            {
                if (!_cachedLibAttribs.HasValue)
                {
                    target_ITypeLib.GetLibAttr(out IntPtr typeLibAttributesPtr);
                    _cachedLibAttribs = StructHelper.ReadStructureUnsafe<ComTypes.TYPELIBATTR>(typeLibAttributesPtr);
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
            if (!HasVBEExtensions)
            {
                throw new InvalidOperationException("This TypeLib does not represent a VBE project, so we cannot compile it");
            }

            try
            {
                target_IVBEProject.CompileProject();
                return true;
            }
            catch (Exception e)
            {
#if DEBUG
                if (e.HResult != (int)KnownComHResults.E_VBA_COMPILEERROR)
                {
                    // this is for debug purposes, to see if the compiler ever returns other errors on failure
                    throw new InvalidOperationException("Unrecognised VBE compiler error: \n" + e.ToString());
                }
#endif
                return false;
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
                if (!HasVBEExtensions)
                {
                    throw new InvalidOperationException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }

                return target_IVBEProject.GetConditionalCompilationArgs();
            }

            set
            {
                if (!HasVBEExtensions)
                {
                    throw new InvalidOperationException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }

                target_IVBEProject.SetConditionalCompilationArgs(value);
            }
        }

        /// <summary>
        /// Exposes the conditional compilation arguments defined in the BA project represented by this ITypeLib
        /// as a dictionary of key/value pairs
        /// </summary>
        public Dictionary<string, short> ConditionalCompilationArguments
        {
            get
            {
                if (!HasVBEExtensions)
                {
                    throw new InvalidOperationException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }

                string args = target_IVBEProject.GetConditionalCompilationArgs();

                if (args.Length > 0)
                {
                    string[] argsArray = args.Split(new[] { ':' });
                    return argsArray.Select(item => item.Split('=')).ToDictionary(s => s[0].Trim(), s => short.Parse(s[1]));
                }
                else
                {
                    return new Dictionary<string, short>();
                }
            }

            set
            {
                if (!HasVBEExtensions)
                {
                    throw new InvalidOperationException("This ITypeLib is not hosted by the VBE, so does not support ConditionalCompilationArguments");
                }

                var rawArgsString = string.Join(" : ", value.Select(x => x.Key + " = " + x.Value));
                ConditionalCompilationArgumentsRaw = rawArgsString;
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
            ((ITypeLib_Ptrs)target_ITypeLib).GetTypeInfoOfGuid(guid, out var typeInfoPtr);
            using (var outVal = new TypeInfoWrapper(typeInfoPtr)) // takes ownership of the COM reference
            {
                ppTInfo = outVal;

                _typeInfosWrapped?.Dispose();
                _typeInfosWrapped = _typeInfosWrapped ?? new DisposableList<TypeInfoWrapper>();
                _typeInfosWrapped.Add(outVal);
            }
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
    public sealed class VBETypeLibsIterator : IEnumerable<TypeLibWrapper>, IEnumerator<TypeLibWrapper>
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

        TypeLibWrapper IEnumerator<TypeLibWrapper>.Current => new TypeLibWrapper(_currentTypeLibPtr, isRefCountedInput: false);
        object IEnumerator.Current => new TypeLibWrapper(_currentTypeLibPtr, isRefCountedInput: false);

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
    public class VBETypeLibsAccessor : DisposableList<TypeLibWrapper>
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

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}
