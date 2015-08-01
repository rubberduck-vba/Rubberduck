using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.Conventions;
using Ninject.Extensions.NamedScope;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.CodeInspections;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Root
{
    // todo: rename to RubberduckModule and derive from NinjectModule
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
            _kernel.Bind<App>().ToSelf();

            // bind VBE and AddIn dependencies to host-provided instances.
            _kernel.Bind<VBE>().ToConstant(vbe);
            _kernel.Bind<AddIn>().ToConstant(addin);

            BindCodeInspectionTypes();

            var assemblies = new[]
            {
                Assembly.GetExecutingAssembly(),
                Assembly.GetAssembly(typeof(IHostApplication)),
                Assembly.GetAssembly(typeof(IRubberduckParser))
            };

            BindMenuTypes();
            BindToolbarTypes();

            ApplyConfigurationConvention(assemblies);
            ApplyDefaultInterfacesConvention(assemblies);
            ApplyAbstractFactoryConvention(assemblies);
        }

        private void BindMenuTypes()
        {
            _kernel.Bind<IMenu>().To<RubberduckMenu>(); // todo: confirm RubberduckMenuFactory is actually needed
            _kernel.Bind<IMenu>().To<FormContextMenu>().WhenTargetHas<FormContextMenuAttribute>();
        }

        private void BindToolbarTypes()
        {
            _kernel.Bind<IToolbar>().To<CodeInspectionsToolbar>().WhenTargetHas<CodeInspectionsToolbarAttribute>();
        }

        private void ApplyDefaultInterfacesConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                // inspections & factories have their own binding rules
                .Where(type => !type.Name.EndsWith("Factory") && !type.GetInterfaces().Contains(typeof(IInspection)))
                .BindDefaultInterface()
                .Configure(binding => binding.InCallScope()));
        }

        // note convention: settings namespace classes are injected in singleton scope
        private void ApplyConfigurationConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                .InNamespaceOf<IGeneralConfigService>()
                .BindAllInterfaces()
                .Configure(binding => binding.InSingletonScope()));
        }

        // note convention: abstract factory interface names end with "Factory".
        private void ApplyAbstractFactoryConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllInterfaces()
                .Where(type => type.Name.EndsWith("Factory"))
                .BindToFactory()
                .Configure(binding => binding.InSingletonScope()));
        }

        // note: IInspection implementations are discovered in the Rubberduck assembly via reflection.
        private void BindCodeInspectionTypes()
        {
            var inspections = Assembly.GetExecutingAssembly()
                                      .GetTypes()
                                      .Where(type => type.GetInterfaces().Contains(typeof (IInspection)));

            // multibinding for IEnumerable<IInspection> dependency
            foreach (var inspection in inspections)
            {
                _kernel.Bind<IInspection>().To(inspection).InSingletonScope();
            }
        }
    }
}
