using Microsoft.Vbe.Interop;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerAddClassModuleCommand : CommandBase
    {
        private readonly VBE _vbe;

        public CodeExplorerAddClassModuleCommand(VBE vbe)
        {
            _vbe = vbe;
        }

        public override bool CanExecute(object parameter)
        {
            return _vbe.ActiveVBProject != null;
        }

        public override void Execute(object parameter)
        {
            _vbe.ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
        }
    }
}