using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Exposes an enumerable collection of TypeInfo objects exposed by this ITypeLib
    /// </summary>
    internal class TypeInfoWrapperCollection : IndexedCollectionBase<ITypeInfoWrapper>, ITypeInfoWrapperCollection
    {
        private readonly ITypeLibWrapper _parent;
        public TypeInfoWrapperCollection(ITypeLibWrapper parent) => _parent = parent;
        public override int Count => _parent.TypesCount;
        public override ITypeInfoWrapper GetItemByIndex(int index)
        {
            var hr = _parent.GetSafeTypeInfoByIndex(index, out var retVal);

            if (ComHelper.HRESULT_FAILED(hr))
            {
                throw new System.Runtime.InteropServices.COMException("TypeInfosCollection::GetItemByIndex failed.", hr);
            }

            return retVal;
        }

        public ITypeInfoWrapper Find(string searchTypeName)
        {
            foreach (var typeInfo in this)
            {
                if (typeInfo.Name == searchTypeName) return typeInfo;
                typeInfo.Dispose();
            }
            return null;
        }

        public ITypeInfoWrapper Get(string searchTypeName)
        {
            var retVal = Find(searchTypeName);
            if (retVal == null)
            {
                throw new ArgumentException($"TypeInfosCollection::Get failed. '{searchTypeName}' component not found.");
            }
            return retVal;
        }
    }
}
