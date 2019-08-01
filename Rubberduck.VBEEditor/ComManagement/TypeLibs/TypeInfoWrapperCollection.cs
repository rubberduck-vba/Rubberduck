using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;

/// <summary>
/// For usage examples, please see VBETypeLibsAPI
/// </summary>
/// <remarks>
/// TypeInfos from a VBA hosted project, and obtained through VBETypeLibsAccessor will have the following behaviours:
/// 
///   will expose both public and private prcoedures and fields
///   will expose constants values, but they are unnamed (their member IDs will be MEMBERID_NIL)
///   enumerations are not exposed directly in the type library
///   enumerations may be referenced by field/argument datatypes, and the ITypeInfos for them are then accessible that way
///   UDTs are not exposed directly in the type library
///   UDTs may be referenced by field/argument datatypes, and as such the ITypeInfos for them are then accessible that way
///   
/// TypeInfos obtained by other means (such as the IDispatch::GetTypeInfo method) usually expose more restricted
/// versions of ITypeInfo which may not expose private members
/// </remarks>

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
