using System.Runtime.InteropServices;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class IndentCurrentProcedureCommand : CommandBase
    {
        private readonly IIndenter _indenter;

        public IndentCurrentProcedureCommand(IIndenter indenter)
        {
            _indenter = indenter;
        }

        public override void Execute(object parameter)
        {
            _indenter.IndentCurrentProcedure();
        }
    }

    [ComVisible(false)]
    public class IndentCurrentModuleCommand : CommandBase
    {
        private readonly IIndenter _indenter;

        public IndentCurrentModuleCommand(IIndenter indenter)
        {
            _indenter = indenter;
        }

        public override void Execute(object parameter)
        {
            _indenter.IndentCurrentModule();
        }
    }
}