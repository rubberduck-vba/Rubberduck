using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Exposes an enumerable collection of variables[fields] provided by the ITypeInfo
    /// </summary>
    public class TypeInfoVariablesCollection : IIndexedCollectionBase<TypeInfoVariable>
    {
        private readonly ComTypes.ITypeInfo _parent;
        private readonly int _count;

        public TypeInfoVariablesCollection(ComTypes.ITypeInfo parent, ComTypes.TYPEATTR attributes)
        {
            _parent = parent;
            _count = attributes.cVars;
        }
        public override int Count => _count;
        public override TypeInfoVariable GetItemByIndex(int index) => new TypeInfoVariable(_parent, index);
    }
}
