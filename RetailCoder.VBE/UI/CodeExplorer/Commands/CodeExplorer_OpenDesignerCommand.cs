using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_OpenDesignerCommand : CommandBase
    {
        public override bool CanExecute(object parameter)
        {
            if (parameter == null)
            {
                return false;   
            }

            try
            {
                var declaration = ((CodeExplorerItemViewModel) parameter).GetSelectedDeclaration();
                return declaration != null && declaration.DeclarationType == DeclarationType.ClassModule &&
                       declaration.QualifiedName.QualifiedModuleName.Component.Designer != null;
            }
            catch (COMException)
            {
                return false;   // component was probably removed
            }
        }

        public override void Execute(object parameter)
        {
            ((ICodeExplorerDeclarationViewModel) parameter).Declaration
                .QualifiedName.QualifiedModuleName.Component.DesignerWindow()
                .Visible = true;
        }
    }
}