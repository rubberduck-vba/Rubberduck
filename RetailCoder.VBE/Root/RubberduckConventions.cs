using System.Linq;
using System.Reflection;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.Conventions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.SourceControl;
using Rubberduck.VBEditor;

namespace Rubberduck.Root
{
    internal class RubberduckConventions
    {
        private readonly IKernel _kernel;

        public RubberduckConventions(IKernel kernel)
        {
            _kernel = kernel;
        }

        public void Apply(VBE vbe, AddIn addin)
        {
            _kernel.Bind<App>().ToSelf();

            _kernel.Bind<VBE>().ToConstant(vbe);
            _kernel.Bind<AddIn>().ToConstant(addin);

            // bind IFoo to Foo across all assemblies:
            _kernel.Bind(t => t.FromThisAssembly().SelectAllClasses().BindDefaultInterface());
            _kernel.Bind(t => t.FromAssemblyContaining<Selection>().SelectAllClasses().BindDefaultInterface());
            _kernel.Bind(t => t.FromAssemblyContaining<IRubberduckParser>().SelectAllClasses().BindDefaultInterface()
                .ConfigureFor<RubberduckParser>(service => service.InSingletonScope()));
        }
    }
}
