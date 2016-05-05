using Microsoft.Vbe.Interop;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerAddUserFormCommand : CommandBase
    {
        private readonly VBE _vbe;

        public CodeExplorerAddUserFormCommand(VBE vbe)
        {
            _vbe = vbe;
        }

        public override bool CanExecute(object parameter)
        {
            return _vbe.ActiveVBProject != null;
        }

        public override void Execute(object parameter)
        {
            _vbe.ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_MSForm);
        }
    }
}