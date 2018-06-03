using System.Collections.Generic;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VB6;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.Root
{
    // Resolves SafeComWrapper providers from raw COM vbe and addin objects.
    // We need these so early that IoC hasn't been set up yet.
    // Using the Provider pattern to avoid outgoing dependencies to COM interop assemblies.
    public static class RootComWrapperFactory
    {
        public static IVBE GetVbeWrapper(object vbeComObject)
        {
            var vbeProviders = new HashSet<ISafeComWrapperProvider<IVBE>>
            {
                new VB6VBEProvider(),
                new VBAVBEProvider()
            };
            var factory = new VBEFactory(vbeProviders);

            return factory.Create(vbeComObject);
        }

        public static IAddIn GetAddInWrapper(object addInComObject)
        {
            var addInProviders = new HashSet<ISafeComWrapperProvider<IAddIn>>
            {
                new VB6AddInProvider(),
                new VBAAddInProvider()
            };
            var factory = new AddInFactory(addInProviders);

            return factory.Create(addInComObject);
        }
    }
}
