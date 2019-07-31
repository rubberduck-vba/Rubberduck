using System.Collections.Generic;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Exposes an enumerable collection of references used by the VBE type library
    /// </summary>
    internal class TypeLibReferenceCollection : IndexedCollectionBase<TypeLibReference>, ITypeLibReferenceCollection
    {
        private readonly TypeLibVBEExtensions _parent;
        public TypeLibReferenceCollection(TypeLibVBEExtensions parent) => _parent = parent;
        public override int Count => _parent.GetVBEReferencesCount();
        public override TypeLibReference GetItemByIndex(int index) => _parent.GetVBEReferenceByIndex(index);

        ITypeLibReference ITypeLibReferenceCollection.GetItemByIndex(int index) => GetItemByIndex(index);
        IEnumerator<ITypeLibReference> ITypeLibReferenceCollection.GetEnumerator() => GetEnumerator();
    }
}
