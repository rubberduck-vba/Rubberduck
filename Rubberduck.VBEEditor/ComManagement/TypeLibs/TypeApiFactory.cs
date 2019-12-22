using System;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.DebugInternal;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;

// TODO The tracers are broken - using them will cause a NRE inside the 
// unmanaged boundary. If we need to enable them for diagnostics, this needs
// to be fixed first. 

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
            ITypeLibWrapper wrapper = new TypeLibWrapper(rawObjectPtr, addRef);
            TraceWrapper(ref wrapper);
            return wrapper;
        }

        [Conditional("TRACE_TYPEAPI")]
        private static void TraceWrapper(ref ITypeLibWrapper wrapper)
        {
            wrapper = new TypeLibWrapperTracer(wrapper, (ITypeLibInternal)wrapper);
        }

        internal static ITypeInfoWrapper GetTypeInfoWrapper(IntPtr rawObjectPtr, int? parentUserFormUniqueId = null)
        {
            ITypeInfoWrapper wrapper = new TypeInfoWrapper(rawObjectPtr, parentUserFormUniqueId);
            TraceWrapper(ref wrapper);
            return wrapper;
        }

        internal static ITypeInfoWrapper GetTypeInfoWrapper(ITypeInfo rawTypeInfo)
        {
            ITypeInfoWrapper wrapper = new TypeInfoWrapper(rawTypeInfo);
            TraceWrapper(ref wrapper);
            return wrapper;
        }

        [Conditional("TRACE_TYPEAPI")]
        private static void TraceWrapper(ref ITypeInfoWrapper wrapper)
        {
            wrapper = new TypeInfoWrapperTracer(wrapper, (ITypeInfoInternal)wrapper);
        }
    }
}
