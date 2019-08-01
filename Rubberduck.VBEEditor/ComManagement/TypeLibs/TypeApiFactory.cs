#define TRACE_TYPEAPI

using System;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;

#if DEBUG && TRACE_TYPEAPI
using Rubberduck.VBEditor.ComManagement.TypeLibs.DebugInternal;
#endif

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    internal static class TypeApiFactory
    {
        internal static ITypeLibWrapper GetTypeLibWrapper(IntPtr rawObjectPtr, bool addRef)
        {
            var wrapper = new TypeLibWrapper(rawObjectPtr, addRef);
#if DEBUG && TRACE_TYPEAPI
            return new TypeLibWrapperTracer(wrapper);
#else
            return wrapper;
#endif
        }

        internal static ITypeInfoWrapper GetTypeInfoWrapper(IntPtr rawObjectPtr, int? parentUserFormUniqueId = null)
        {
            var wrapper = new TypeInfoWrapper(rawObjectPtr, parentUserFormUniqueId);
#if DEBUG && TRACE_TYPEAPI
            return new TypeInfoWrapperTracer(wrapper);
#else
            return wrapper;
#endif
        }

        internal static ITypeInfoWrapper GetTypeInfoWrapper(ITypeInfo rawTypeInfo)
        {
            var wrapper = new TypeInfoWrapper(rawTypeInfo);
#if DEBUG && TRACE_TYPEAPI
            return new TypeInfoWrapperTracer(wrapper);
#else
            return wrapper;
#endif
        }
    }
}
