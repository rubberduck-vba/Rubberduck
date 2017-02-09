using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;

namespace Rubberduck.Parsing.ComReflection
{
    [DebuggerDisplay("{Name}")]
    public class ComProject : ComBase
    {
        public static readonly ConcurrentDictionary<Guid, ComType> KnownTypes = new ConcurrentDictionary<Guid, ComType>();
        public static readonly ConcurrentDictionary<Guid, ComEnumeration> KnownEnumerations = new ConcurrentDictionary<Guid, ComEnumeration>(); 

        public string Path { get; set; }
        public long MajorVersion { get; private set; }
        public long MinorVersion { get; private set; }

        // YGNI...
        // ReSharper disable once NotAccessedField.Local
        private TypeLibTypeFlags _flags;

        private readonly List<ComAlias> _aliases = new List<ComAlias>();
        public IEnumerable<ComAlias> Aliases
        {
            get { return _aliases; }
        }

        private readonly List<ComInterface> _interfaces = new List<ComInterface>();
        public IEnumerable<ComInterface> Interfaces
        {
            get { return _interfaces; }
        }

        private readonly List<ComEnumeration> _enumerations = new List<ComEnumeration>();
        public IEnumerable<ComEnumeration> Enumerations
        {
            get { return _enumerations; }
        }

        private readonly List<ComCoClass> _classes = new List<ComCoClass>();
        public IEnumerable<ComCoClass> CoClasses
        {
            get { return _classes; }
        }

        private readonly List<ComModule> _modules = new List<ComModule>();
        public IEnumerable<ComModule> Modules
        {
            get { return _modules; }
        }

        private readonly List<ComStruct> _structs = new List<ComStruct>();
        public IEnumerable<ComStruct> Structs
        {
            get { return _structs; }
        }

        public IEnumerable<IComType> Members
        {
            get
            {
                //Note - Enums and Types should enumerate *last*. That will prevent a duplicate module in the unlikely(?)
                //instance where the TypeLib defines a module named "Enums" or "Types".
                return _modules.Cast<IComType>()
                    .Union(_interfaces)
                    .Union(_classes)
                    .Union(_enumerations)
                    .Union(_structs);
            }
        } 

        public ComProject(ITypeLib typeLibrary) : base(typeLibrary, -1)
        {   
            ProcessLibraryAttributes(typeLibrary);
            LoadModules(typeLibrary);
        }

        private void ProcessLibraryAttributes(ITypeLib typeLibrary)
        {
            try
            {
                IntPtr attribPtr;
                typeLibrary.GetLibAttr(out attribPtr);
                var typeAttr = (TYPELIBATTR)Marshal.PtrToStructure(attribPtr, typeof(TYPELIBATTR));

                MajorVersion = typeAttr.wMajorVerNum;
                MinorVersion = typeAttr.wMinorVerNum;
                _flags = (TypeLibTypeFlags)typeAttr.wLibFlags;
                Guid = typeAttr.guid;
            }
            catch (COMException) { }
        }

        private void LoadModules(ITypeLib typeLibrary)
        {
            var typeCount = typeLibrary.GetTypeInfoCount();
            for (var index = 0; index < typeCount; index++)
            {                
                try
                {
                    ITypeInfo info;
                    typeLibrary.GetTypeInfo(index, out info);
                    IntPtr typeAttributesPointer;
                    info.GetTypeAttr(out typeAttributesPointer);
                    var typeAttributes = (TYPEATTR)Marshal.PtrToStructure(typeAttributesPointer, typeof(TYPEATTR));

                    ComType type;
                    KnownTypes.TryGetValue(typeAttributes.guid, out type);

                    switch (typeAttributes.typekind)
                    {
                        case TYPEKIND.TKIND_ENUM:
                            var enumeration = type ?? new ComEnumeration(typeLibrary, info, typeAttributes, index);
                            _enumerations.Add(enumeration as ComEnumeration);
                            if (type != null) KnownTypes.TryAdd(typeAttributes.guid, enumeration);
                            break;
                        case TYPEKIND.TKIND_COCLASS:
                            var coclass = type ?? new ComCoClass(typeLibrary, info, typeAttributes, index);
                            _classes.Add(coclass as ComCoClass);
                            if (type != null) KnownTypes.TryAdd(typeAttributes.guid, coclass);
                            break;
                        case TYPEKIND.TKIND_DISPATCH:
                        case TYPEKIND.TKIND_INTERFACE:
                            var intface = type ?? new ComInterface(typeLibrary, info, typeAttributes, index);
                            _interfaces.Add(intface as ComInterface);
                            if (type != null) KnownTypes.TryAdd(typeAttributes.guid, intface);
                            break;
                        case TYPEKIND.TKIND_RECORD:
                            var structure = new ComStruct(typeLibrary, info, typeAttributes, index);
                            _structs.Add(structure);
                            break;
                        case TYPEKIND.TKIND_MODULE:
                            var module = type ?? new ComModule(typeLibrary, info, typeAttributes, index);
                            _modules.Add(module as ComModule);
                            if (type != null) KnownTypes.TryAdd(typeAttributes.guid, module);
                            break;
                        case TYPEKIND.TKIND_ALIAS:
                            var alias = new ComAlias(typeLibrary, info, index, typeAttributes);
                            _aliases.Add(alias);
                            break;
                        case TYPEKIND.TKIND_UNION:
                            //TKIND_UNION is not a supported member type in VBA.
                            break;
                        default:
                            throw new NotImplementedException(string.Format("Didn't expect a TYPEATTR with multiple typekind flags set in {0}.", Path));
                    }
                    info.ReleaseTypeAttr(typeAttributesPointer);
                }
                catch (COMException) { }
            }
            ApplySpecificLibraryTweaks();
        }

        private void ApplySpecificLibraryTweaks()
        {
            if (!Name.ToUpper().Equals("EXCEL")) return;
            var application = _classes.SingleOrDefault(x => x.Guid.ToString().Equals("00024500-0000-0000-c000-000000000046"));
            var worksheetFunction = _interfaces.SingleOrDefault(i => i.Guid.ToString().Equals("00020845-0000-0000-c000-000000000046"));
            if (application != null && worksheetFunction != null)
            {
                application.AddInterface(worksheetFunction);
            }
        }
    }
}
