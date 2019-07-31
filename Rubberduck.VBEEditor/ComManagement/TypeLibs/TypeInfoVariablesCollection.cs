using System.Collections.Generic;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Exposes an enumerable collection of variables[fields] provided by the ITypeInfo
    /// </summary>
    internal class TypeInfoVariablesCollection : IndexedCollectionBase<TypeInfoVariable>, ITypeInfoVariablesCollection
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

        ITypeInfoVariable ITypeInfoVariablesCollection.GetItemByIndex(int index)
        {
            return GetItemByIndex(index);
        }

        IEnumerator<ITypeInfoVariable> ITypeInfoVariablesCollection.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
