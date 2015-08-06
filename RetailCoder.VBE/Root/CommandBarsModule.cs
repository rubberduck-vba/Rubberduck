using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ninject;
using Ninject.Modules;

namespace Rubberduck.Root
{
    public class CommandBarsModule : NinjectModule
    {
        private readonly IKernel _kernel;

        public CommandBarsModule(IKernel kernel)
        {
            _kernel = kernel;
        }

        public override void Load()
        {
            throw new NotImplementedException();
        }
    }
}
