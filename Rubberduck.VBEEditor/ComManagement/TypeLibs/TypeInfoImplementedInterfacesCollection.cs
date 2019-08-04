using System;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Exposes an enumerable collection of implemented interfaces provided by the <see cref="ComTypes.ITypeInfo"/>
    /// </summary>
    internal class TypeInfoImplementedInterfacesCollection : IndexedCollectionBase<ITypeInfoWrapper>, ITypeInfoImplementedInterfacesCollection
    {
        private readonly ComTypes.ITypeInfo _parent;
        private readonly int _count;
        public TypeInfoImplementedInterfacesCollection(ComTypes.ITypeInfo parent, ComTypes.TYPEATTR attributes)
        {
            _parent = parent;
            _count = attributes.cImplTypes;
        }
        public override int Count => _count;
        public override ITypeInfoWrapper GetItemByIndex(int index)
        {
            _parent.GetRefTypeOfImplType(index, out var href);
            _parent.GetRefTypeInfo(href, out var ti);

            return TypeApiFactory.GetTypeInfoWrapper(ti);
        }

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
                    if ((typeInfo.ContainerName == containerName) && (typeInfo.Name == interfaceName)) return true;
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

        public ITypeInfoWrapper Get(string searchTypeName)
        {
            foreach (var typeInfo in this)
            {
                if (typeInfo.Name == searchTypeName) return typeInfo;
                typeInfo.Dispose();
            }

            throw new ArgumentException($"TypeInfoWrapper::Get failed. '{searchTypeName}' component not found.");
        }
    }
}
