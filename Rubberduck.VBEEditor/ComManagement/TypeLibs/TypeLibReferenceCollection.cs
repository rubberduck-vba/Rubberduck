namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Exposes an enumerable collection of references used by the VBE type library
    /// </summary>
    public class TypeLibReferenceCollection : IIndexedCollectionBase<TypeLibReference>
    {
        private readonly TypeLibVBEExtensions _parent;
        public TypeLibReferenceCollection(TypeLibVBEExtensions parent) => _parent = parent;
        public override int Count => _parent.GetVBEReferencesCount();
        public override TypeLibReference GetItemByIndex(int index) => _parent.GetVBEReferenceByIndex(index);
    }
}
