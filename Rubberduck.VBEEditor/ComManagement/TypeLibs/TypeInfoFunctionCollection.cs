using System.Collections.Generic;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Exposes an enumerable collection of functions provided by the ITypeInfo
    /// </summary>
    internal class TypeInfoFunctionCollection : IndexedCollectionBase<TypeInfoFunction>, ITypeInfoFunctionCollection
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

        public TypeInfoFunction Find(string name, PROCKIND procKind)
        {
            foreach (var func in this)
            {
                if ((func.Name == name) && (func.ProcKind == procKind)) return func;
            }
            return null;
        }

        ITypeInfoFunction ITypeInfoFunctionCollection.GetItemByIndex(int index)
        {
            return GetItemByIndex(index);
        }

        ITypeInfoFunction ITypeInfoFunctionCollection.Find(string name, PROCKIND procKind)
        {
            return Find(name, procKind);
        }

        IEnumerator<ITypeInfoFunction> ITypeInfoFunctionCollection.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
