using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Command
{
    public interface ICommand
    {
        void Execute();
    }

    public interface ICommand<in T>
    {
        void Execute(T parameter);
    }
}
