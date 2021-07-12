using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Com.Extensions
{
    /// <summary>
    /// Type API specific extensions for managing memory allocated by calls to Get*** which must be
    /// followed by a corresponding Release ***. This is used by other extensions classes:
    /// <see cref="TypeLibExtensions"/>
    /// <see cref="TypeInfoExtensions"/>
    /// </summary>
    internal static class ExtensionHelper
    {
        internal static void UsingPtrToStructure<T>(Func<IntPtr, IntPtr> acquireBlock, Action<T> usingBlock, Action<IntPtr> releaseBlock)
        {
            var ptr = IntPtr.Zero;
            try
            {
                ptr = acquireBlock.Invoke(ptr);
                T t = default;
                if (ptr != IntPtr.Zero)
                {
                    t = Marshal.PtrToStructure<T>(ptr);
                }
                usingBlock.Invoke(t);
            }
            finally
            {
                if (ptr != IntPtr.Zero)
                {
                    releaseBlock.Invoke(ptr);
                }
            }
        }
    }
}
