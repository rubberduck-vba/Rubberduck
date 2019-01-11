using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Exposes an enumerable collection of functions provided by the ITypeInfo
    /// </summary>
    public class TypeInfoFunctionCollection : IIndexedCollectionBase<TypeInfoFunction>
    {
        private readonly ComTypes.ITypeInfo _parent;
        private readonly int _count;

        public TypeInfoFunctionCollection(ComTypes.ITypeInfo parent, ComTypes.TYPEATTR attributes)
        {
            _parent = parent;
            _count = attributes.cFuncs;
        }

        public override int Count => _count;
        public override TypeInfoFunction GetItemByIndex(int index) => new TypeInfoFunction(_parent, index);

        public TypeInfoFunction Find(string name, TypeInfoFunction.PROCKIND procKind)
        {
            foreach (var func in this)
            {
                if ((func.Name == name) && (func.ProcKind == procKind)) return func;
            }
            return null;
        }
    }
}
