using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.VbeRuntime;

namespace Rubberduck
{
    /// <summary>
    /// ANTI-PATTERN: Service locator for COM class. Think carefully before using it, and regret it.
    /// </summary>
    /// <remarks>
    /// This is a hacky workaround to provide support to COM-visible classes without breaking the
    /// interface or violating the security settings of the Office host. Because a COM class must
    /// have a parameterless constructor if it is to be newed up and because COM class cannot come
    /// from the IoC container nor have any dependencies coming out of it, we use the service
    /// locator anti-pattern here to provide the necessary functionality for the COM classes'
    /// internal implementations. The use should never expand beyond that limited scope. 
    /// </remarks>
    internal static class VbeProvider
    {
        internal static void Initialize(IVBE vbe)
        {
            Vbe = vbe;
            VbeRuntime = new VbeNativeApiAccessor(vbe);
        }

        internal static void Terminate()
        {
            Vbe = null;
            VbeRuntime = null;
        }

        internal static IVBE Vbe { get; private set; }
        internal static IVbeNativeApi VbeRuntime { get; private set; }
    }
}
