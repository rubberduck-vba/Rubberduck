using Castle.Core;
using Castle.MicroKernel.Context;
using Castle.MicroKernel.Handlers;
using System;
using System.Linq;

namespace Rubberduck.Root
{
    class FixedGenericAppender : IGenericImplementationMatchingStrategy
    {
        private readonly Type[] closingGenerics;

        public FixedGenericAppender(Type[] closingGenerics)
        {
            this.closingGenerics = closingGenerics;
        }

        public Type[] GetGenericArguments(ComponentModel model, CreationContext context)
        {
            return context.GenericArguments.Union(closingGenerics).ToArray();
        }
    }
}
