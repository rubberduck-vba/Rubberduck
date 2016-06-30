using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_OpenProjectPropertiesCommand : CommandBase
    {
        private readonly VBE _vbe;

        public CodeExplorer_OpenProjectPropertiesCommand(VBE vbe) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
        }

        protected override void ExecuteImpl(object parameter)
        {
            const int openProjectPropertiesId = 2578;

            _vbe.CommandBars.FindControl(Id: openProjectPropertiesId).Execute();
        }
    }
}
