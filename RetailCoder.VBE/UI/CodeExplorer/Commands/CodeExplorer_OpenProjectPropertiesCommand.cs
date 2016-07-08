using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
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

        protected override bool CanExecuteImpl(object parameter)
        {
            return parameter != null || _vbe.VBProjects.Count == 1;
        }

        protected override void ExecuteImpl(object parameter)
        {
            const int openProjectPropertiesId = 2578;

            if (_vbe.VBProjects.Count == 1)
            {
                _vbe.CommandBars.FindControl(Id: openProjectPropertiesId).Execute();
                return;
            }

            var node = parameter as CodeExplorerItemViewModel;
            while (!(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent; // the project node is an ICodeExplorerDeclarationViewModel--no worries here
            }

            try
            {
                _vbe.ActiveVBProject = node.GetSelectedDeclaration().Project;
            }
            catch (COMException)
            {
                return; // the project was probably removed from the VBE, but not from the CE
            }

            _vbe.CommandBars.FindControl(Id: openProjectPropertiesId).Execute();
        }
    }
}
