using Microsoft.Vbe.Interop;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerAddStdModuleCommand : CommandBase
    {
        private readonly VBE _vbe;

        public CodeExplorerAddStdModuleCommand(VBE vbe)
        {
            _vbe = vbe;
        }

        public override bool CanExecute(object parameter)
        {
            return _vbe.ActiveVBProject != null;
        }

        public override void Execute(object parameter)
        {
            _vbe.ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
        }
    }
}