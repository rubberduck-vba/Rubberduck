using Castle.Core;
using Castle.MicroKernel;
using Castle.MicroKernel.Context;
using Moq;

namespace RubberduckTests.Refactoring.MockIoC
{
    internal class AutoMoqResolver : ISubDependencyResolver
    {
        private readonly IKernel _kernel;

        public AutoMoqResolver(IKernel kernel)
        {
            _kernel = kernel;
        }

        public bool CanResolve(
            CreationContext context,
            ISubDependencyResolver contextHandlerResolver,
            ComponentModel model,
            DependencyModel dependency)
        {
            return dependency.TargetType.Namespace.StartsWith("Rubberduck")
                   && !dependency.TargetType.Name.EndsWith("Factory")
                   && dependency.TargetType.IsInterface;
        }

        public object Resolve(
            CreationContext context,
            ISubDependencyResolver contextHandlerResolver,
            ComponentModel model,
            DependencyModel dependency)
        {
            var mockType = typeof(Mock<>).MakeGenericType(dependency.TargetType);
            return ((Mock)_kernel.Resolve(mockType)).Object;
        }
    }
}
