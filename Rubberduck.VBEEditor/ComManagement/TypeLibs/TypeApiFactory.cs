using System;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;

// TODO The tracers are broken - using them will cause a NRE inside the 
// unmanaged boundary. If we need to enable them for diagnostics, this needs
// to be fixed first. 
#if DEBUG && TRACE_TYPEAPI
using Rubberduck.VBEditor.ComManagement.TypeLibs.DebugInternal;
#endif

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Abstracts out the creation of the custom implementations of
    /// <see cref="ITypeLib"/> and <see cref="ITypeInfo"/>, mainly to
    /// make it easier to compose the implementation. For example, tracing
    /// can be enabled via the factory with appropriate compilation flags. 
    /// </summary>
    internal static class TypeApiFactory
    {
        internal static ITypeLibWrapper GetTypeLibWrapper(IntPtr rawObjectPtr, bool addRef)
        {
            var wrapper = new TypeLibWrapper(rawObjectPtr, addRef);
#if DEBUG && TRACE_TYPEAPI
            return new TypeLibWrapperTracer(wrapper, wrapper);
#else
            return wrapper;
#endif
        }

        internal static ITypeInfoWrapper GetTypeInfoWrapper(IntPtr rawObjectPtr, int? parentUserFormUniqueId = null)
        {
            var wrapper = new TypeInfoWrapper(rawObjectPtr, parentUserFormUniqueId);
#if DEBUG && TRACE_TYPEAPI
            return new TypeInfoWrapperTracer(wrapper, wrapper);
#else
            return wrapper;
#endif
        }

        internal static ITypeInfoWrapper GetTypeInfoWrapper(ITypeInfo rawTypeInfo)
        {
            var wrapper = new TypeInfoWrapper(rawTypeInfo);
#if DEBUG && TRACE_TYPEAPI
            return new TypeInfoWrapperTracer(wrapper, wrapper);
#else
            return wrapper;
#endif
        }
    }
}
