using System;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor
{
    // Factory for addin providers.
    // Implemented using Lazy on the provider function to guard for any race conditions occurring during construction.
    public class AddInFactory
    {
        private readonly Lazy<Func<object, IAddIn>> _provide;

        public AddInFactory(IEnumerable<ISafeComWrapperProvider<IAddIn>> vbeProviders)
        {
            _provide = new Lazy<Func<object, IAddIn>>(() => vbeComObject =>
            {
                foreach (var vbeProvider in vbeProviders)
                {
                    if (vbeProvider.CanProvideFor(vbeComObject))
                    {
                        return vbeProvider.Provide(vbeComObject);
                    }
                }
                throw new NotSupportedException($"Add-in type {vbeComObject.GetType().Name} is not supported");
            });

        }

        public IAddIn Create(object vbeCom)
        {
            return _provide.Value(vbeCom);
        }
    }
}