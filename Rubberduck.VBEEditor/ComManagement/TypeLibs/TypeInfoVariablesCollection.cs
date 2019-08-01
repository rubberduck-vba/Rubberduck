using System.Collections.Generic;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Exposes an enumerable collection of variables[fields] provided by the ITypeInfo
    /// </summary>
    internal class TypeInfoVariablesCollection : IndexedCollectionBase<ITypeInfoVariable>, ITypeInfoVariablesCollection
    {
        private readonly ComTypes.ITypeInfo _parent;
        private readonly int _count;

        public TypeInfoVariablesCollection(ComTypes.ITypeInfo parent, ComTypes.TYPEATTR attributes)
        {
            _parent = parent;
            _count = attributes.cVars;
        }
        public override int Count => _count;
        
        public override ITypeInfoVariable GetItemByIndex(int index) => new TypeInfoVariable(_parent, index);
    }
}
