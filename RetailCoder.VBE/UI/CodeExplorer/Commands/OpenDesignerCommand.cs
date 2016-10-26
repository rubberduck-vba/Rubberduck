using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class OpenDesignerCommand : CommandBase
    {
        public OpenDesignerCommand() : base(LogManager.GetCurrentClassLogger()) { }

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
                       declaration.QualifiedName.QualifiedModuleName.Component.HasDesigner;
            }
            catch (COMException)
            {
                // thrown when the component reference is stale
                return false;
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            var designer = ((ICodeExplorerDeclarationViewModel) parameter).Declaration.QualifiedName.QualifiedModuleName.Component.DesignerWindow();
            {
                if (!designer.IsWrappingNullReference)
                {
                    designer.IsVisible = true;
                }
            }
        }
    }
}
