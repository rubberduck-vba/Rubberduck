using System;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeInfoImplementedInterfacesCollection
    {
        int Count { get; }
        ITypeInfoWrapper GetItemByIndex(int index);

        /// <summary>
        /// Determines whether the type implements one of the specified interfaces
        /// </summary>
        /// <param name="interfaceProgIds">Array of interface identifiers in the format "LibName.InterfaceName"</param>
        /// <param name="matchedIndex">on return, contains the index into interfaceProgIds that matched, or -1 </param>
        /// <returns>true if the type does implement one of the specified interfaces</returns>
        bool DoesImplement(string[] interfaceProgIds, out int matchedIndex);

        /// <summary>
        /// Determines whether the type implements the specified interface
        /// </summary>
        /// <param name="interfaceProgId">Interface identifier in the format "LibName.InterfaceName"</param>
        /// <returns>true if the type does implement the specified interface</returns>
        bool DoesImplement(string interfaceProgId);

        /// <summary>
        /// Determines whether the type implements the specified interface
        /// </summary>
        /// <param name="containerName">The library container name</param>
        /// <param name="interfaceName">The interface name</param>
        /// <returns>true if the type does implement the specified interface</returns>
        bool DoesImplement(string containerName, string interfaceName);

        /// <summary>
        /// Determines whether the type implements one of the specified interfaces
        /// </summary>
        /// <param name="interfaceIIDs">Array of interface IIDs to match</param>
        /// <param name="matchedIndex">on return, contains the index into interfaceIIDs that matched, or -1 </param>
        /// <returns>true if the type does implement one of the specified interfaces</returns>
        bool DoesImplement(Guid[] interfaceIIDs, out int matchedIndex);

        /// <summary>
        /// Determines whether the type implements the specified interface
        /// </summary>
        /// <param name="interfaceIID">The interface IID to match</param>
        /// <returns>true if the type does implement the specified interface</returns>
        bool DoesImplement(Guid interfaceIID);

        ITypeInfoWrapper Get(string searchTypeName);
        IEnumerator<ITypeInfoWrapper> GetEnumerator();
    }
}