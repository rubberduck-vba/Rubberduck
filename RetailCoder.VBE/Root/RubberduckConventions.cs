using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.Conventions;
using Ninject.Extensions.Factory;
using Ninject.Extensions.NamedScope;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Root
{
    public class RubberduckConventions
    {
        private readonly IKernel _kernel;

        public RubberduckConventions(IKernel kernel)
        {
            _kernel = kernel;
        }

        /// <summary>
        /// Configures the IoC <see cref="IKernel"/>.
        /// </summary>
        /// <param name="vbe">The <see cref="VBE"/> instance provided by the host application.</param>
        /// <param name="addin">The <see cref="AddIn"/> instance provided by the host application.</param>
        public void Apply(VBE vbe, AddIn addin)
        {
            // use a named scope to ensure proper disposal
            var appScopeName = "AppScope";
            _kernel.Bind<App>().ToSelf().DefinesNamedScope(appScopeName);

            // bind VBE and AddIn dependencies to host-provided instances.
            _kernel.Bind<VBE>().ToConstant(vbe).InNamedScope(appScopeName);
            _kernel.Bind<AddIn>().ToConstant(addin).InNamedScope(appScopeName);

            // multi-binding for code inspections:
            var inspections = FindInspectionTypes();
            foreach (var inspection in inspections)
            {
                _kernel.Bind<IInspection>().To(inspection).InSingletonScope();
            }

            var assemblies = new[]
            {
                Assembly.GetExecutingAssembly(),
                Assembly.GetAssembly(typeof(IHostApplication)),
                Assembly.GetAssembly(typeof(IRubberduckParser))
            };

            // note convention: IFoo binds to Foo.
            ApplyDefaultInterfaceConvention(assemblies, appScopeName);

            // note convention: abstract factory interface names end with "Factory".
            ApplyAbstractFactoryConvention(assemblies, appScopeName);
        }

        private void ApplyDefaultInterfaceConvention(IEnumerable<Assembly> assemblies, string appScopeName)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                .Where(type => !type.Name.EndsWith("Factory")) // skip concrete factory types
                .BindDefaultInterface()
                .Configure(binding => binding.InNamedScope(appScopeName)));
        }

        private void ApplyAbstractFactoryConvention(IEnumerable<Assembly> assemblies, string appScopeName)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllInterfaces()
                .Where(type => type.Name.EndsWith("Factory"))
                .BindToFactory()
                .Configure(binding => binding.InNamedScope(appScopeName)));
        }

        private static IEnumerable<Type> FindInspectionTypes()
        {
            return Assembly.GetExecutingAssembly()
                           .GetTypes()
                           .Where(type => type.GetInterfaces().Contains(typeof (IInspection)));
        }
    }
}
