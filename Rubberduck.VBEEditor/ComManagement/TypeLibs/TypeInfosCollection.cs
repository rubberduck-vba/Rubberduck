using System;
using Rubberduck.VBEditor.ComManagement.TypeLibsSupport;

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
    public class TypeInfosCollection : IIndexedCollectionBase<TypeInfoWrapper>
    {
        private readonly TypeLibWrapper _parent;
        public TypeInfosCollection(TypeLibWrapper parent) => _parent = parent;
        public override int Count => _parent.TypesCount;
        public override TypeInfoWrapper GetItemByIndex(int index) => _parent.GetSafeTypeInfoByIndex(index);

        public TypeInfoWrapper Find(string searchTypeName)
        {
            foreach (var typeInfo in this)
            {
                if (typeInfo.Name == searchTypeName) return typeInfo;
                typeInfo.Dispose();
            }
            return null;
        }

        public TypeInfoWrapper Get(string searchTypeName)
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
