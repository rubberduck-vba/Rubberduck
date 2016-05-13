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
                var declaration = GetSelectedDeclaration((CodeExplorerItemViewModel) parameter);
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