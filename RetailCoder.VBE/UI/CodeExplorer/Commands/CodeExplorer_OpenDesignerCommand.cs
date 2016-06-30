using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_OpenDesignerCommand : CommandBase
    {
        public CodeExplorer_OpenDesignerCommand() : base(LogManager.GetCurrentClassLogger()) { }

        protected override bool CanExecuteImpl(object parameter)
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

        protected override void ExecuteImpl(object parameter)
        {
            ((ICodeExplorerDeclarationViewModel) parameter).Declaration
                .QualifiedName.QualifiedModuleName.Component.DesignerWindow()
                .Visible = true;
        }
    }
}
