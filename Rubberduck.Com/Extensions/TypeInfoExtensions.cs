using System;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.Com.Extensions
{
    public static class TypeInfoExtensions
    {
        public static void UsingTypeAttr(this ITypeInfo typeInfo, Action<TYPEATTR> action)
        {
            ExtensionHelper.UsingPtrToStructure(ptr => { typeInfo.GetTypeAttr(out ptr); return ptr; }, action, typeInfo.ReleaseTypeAttr);
        }

        public static void UsingVarDesc(this ITypeInfo typeInfo, int index, Action<VARDESC> action)
        {
            ExtensionHelper.UsingPtrToStructure(ptr => { typeInfo.GetVarDesc(index, out ptr); return ptr; }, action, typeInfo.ReleaseVarDesc);
        }

        public static void UsingFuncDesc(this ITypeInfo typeInfo, int index, Action<FUNCDESC> action)
        {
            ExtensionHelper.UsingPtrToStructure(ptr => { typeInfo.GetFuncDesc(index, out ptr); return ptr; }, action, typeInfo.ReleaseFuncDesc);
        }
    }
}
