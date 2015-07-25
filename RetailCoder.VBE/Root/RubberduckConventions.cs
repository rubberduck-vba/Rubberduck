using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.Factory;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;

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
            // bind VBE and AddIn dependencies to host-provided instances.
            _kernel.Bind<VBE>().ToConstant(vbe);
            _kernel.Bind<AddIn>().ToConstant(addin);
        }
    }
}
