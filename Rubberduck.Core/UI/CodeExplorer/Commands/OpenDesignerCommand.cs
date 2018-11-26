using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class OpenDesignerCommand : CommandBase
    {
        private readonly IProjectsProvider _projectsProvider;

        public OpenDesignerCommand(IProjectsProvider projectsProvider)
            : base(LogManager.GetCurrentClassLogger())
        {
            _projectsProvider = projectsProvider;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (parameter == null)
            {
                return false;   
            }

            try
            {
                var declaration = ((CodeExplorerItemViewModel) parameter).GetSelectedDeclaration();
                return declaration != null && declaration.DeclarationType == DeclarationType.ClassModule &&
                       _projectsProvider.Component(declaration.QualifiedName.QualifiedModuleName).HasDesigner;
            }
            catch (COMException)
            {
                // thrown when the component reference is stale
                return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            var component = _projectsProvider.Component(((ICodeExplorerDeclarationViewModel)parameter).Declaration.QualifiedName.QualifiedModuleName);
            using (var designer = component.DesignerWindow())
            {
                if (!designer.IsWrappingNullReference)
                {
                    designer.IsVisible = true;
                }
            }
        }
    }
}
