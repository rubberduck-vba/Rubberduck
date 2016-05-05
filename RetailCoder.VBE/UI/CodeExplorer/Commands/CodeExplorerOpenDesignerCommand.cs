using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerOpenDesignerCommand : CommandBase
    {
        public override bool CanExecute(object parameter)
        {
            var declaration = GetSelectedDeclaration((CodeExplorerItemViewModel)parameter);
            return declaration != null && declaration.DeclarationType == DeclarationType.ClassModule &&
                    declaration.QualifiedName.QualifiedModuleName.Component.Designer != null;
        }

        public override void Execute(object parameter)
        {
            GetSelectedDeclaration((CodeExplorerItemViewModel) parameter)
                .QualifiedName.QualifiedModuleName.Component.DesignerWindow()
                .Visible = true;
        }

        private Declaration GetSelectedDeclaration(CodeExplorerItemViewModel node)
        {
            if (node is CodeExplorerProjectViewModel)
            {
                return ((CodeExplorerProjectViewModel)node).Declaration;
            }

            if (node is CodeExplorerComponentViewModel)
            {
                return ((CodeExplorerComponentViewModel)node).Declaration;
            }

            if (node is CodeExplorerMemberViewModel)
            {
                return ((CodeExplorerMemberViewModel)node).Declaration;
            }

            return null;
        }
    }
}