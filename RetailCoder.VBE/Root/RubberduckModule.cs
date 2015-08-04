using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.Conventions;
using Ninject.Extensions.NamedScope;
using Ninject.Modules;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.Commands;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Root
{
    public class RubberduckModule : NinjectModule
    {
        private readonly IKernel _kernel;
        private readonly VBE _vbe;
        private readonly AddIn _addin;

        public RubberduckModule(IKernel kernel, VBE vbe, AddIn addin)
        {
            _kernel = kernel;
            _vbe = vbe;
            _addin = addin;
        }

        public override void Load()
        {
            _kernel.Bind<App>().ToSelf();

            // bind VBE and AddIn dependencies to host-provided instances.
            _kernel.Bind<VBE>().ToConstant(_vbe);
            _kernel.Bind<AddIn>().ToConstant(_addin);

            BindCodeInspectionTypes();

            var assemblies = new[]
            {
                Assembly.GetExecutingAssembly(),
                Assembly.GetAssembly(typeof(IHostApplication)),
                Assembly.GetAssembly(typeof(IRubberduckParser))
            };

            BindRubberduckMenu();

            ApplyConfigurationConvention(assemblies);
            ApplyDefaultInterfacesConvention(assemblies);
            ApplyAbstractFactoryConvention(assemblies);
        }

        private void BindRubberduckMenu()
        {
            const int windowMenuId = 30009;
            var menuBarControls = _vbe.CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls, windowMenuId);

            _kernel.Bind(t => t.FromThisAssembly()
                .SelectAllClasses()
                .InNamespaceOf<ICommand>()
                .EndingWith("CommandMenuItem")
                .BindToSelf());

            //_kernel.Bind(t => t.FromThisAssembly()
            //    .SelectAllClasses()
            //    .InNamespaceOf<ICommand>()
            //    .EndingWith("Command")
            //    .Where(type => type.GetInterfaces().Contains(typeof (ICommand)))
            //    .BindAllInterfaces()
            //    .Configure(binding => binding
            //        .When(request => request.Service == typeof(ICommand) 
            //                      && request.Target.Member.DeclaringType.Name.StartsWith("????"))));

            _kernel.Bind<ICommand>().To<AboutCommand>().WhenInjectedExactlyInto<AboutCommandMenuItem>();
            _kernel.Bind<ICommand>().To<OptionsCommand>().WhenInjectedExactlyInto<OptionsCommandMenuItem>();
            _kernel.Bind<ICommand>().To<CodeExplorerCommand>().WhenInjectedExactlyInto<CodeExplorerCommandMenuItem>();

            _kernel.Bind<RubberduckParentMenu>().ToSelf()
                .WithConstructorArgument("parent", menuBarControls)
                .WithConstructorArgument("beforeIndex", beforeIndex);
        }

        private int FindMenuInsertionIndex(CommandBarControls controls, int beforeId)
        {
            for (var i = 1; i <= controls.Count; i++)
            {
                if (controls[i].BuiltIn && controls[i].Id == beforeId)
                {
                    return i;
                }
            }

            return controls.Count;
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

        // note: settings namespace classes are injected in singleton scope
        private void ApplyConfigurationConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                .InNamespaceOf<Configuration>()
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
