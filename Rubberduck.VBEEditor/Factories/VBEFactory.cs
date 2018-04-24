using System;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor
{
    // Factory for VBE providers.
    // Implemented using Lazy on the provider function to guard for any race conditions occurring during construction.
    public class VBEFactory
    {
        private readonly Lazy<Func<object, IVBE>> _provide;

        public VBEFactory(IEnumerable<ISafeComWrapperProvider<IVBE>> vbeProviders)
        {
            _provide = new Lazy<Func<object, IVBE>>(() => vbeComObject =>
            {
                foreach (var vbeProvider in vbeProviders)
                {
                    if (vbeProvider.CanProvideFor(vbeComObject))
                    {
                        return vbeProvider.Provide(vbeComObject);
                    }
                }
                throw new NotSupportedException("Host application not supported");                
            });
            
        }

        public IVBE Create(object vbeCom)
        {
            return _provide.Value(vbeCom);
        }
    }
}
