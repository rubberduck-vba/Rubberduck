using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class OpenDesignerCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly IProjectsProvider _projectsProvider;

        public OpenDesignerCommand(IProjectsProvider projectsProvider)
        {
            _projectsProvider = projectsProvider;
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (!base.EvaluateCanExecute(parameter) || !(parameter is CodeExplorerItemViewModel node))
            {
                return false;   
            }

            try
            {
                var declaration = node.Declaration;
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
            if (!base.EvaluateCanExecute(parameter) || !(parameter is CodeExplorerItemViewModel node) || node.Declaration == null)
            {
                return;
            }

            var component = _projectsProvider.Component(node.Declaration.QualifiedName.QualifiedModuleName);
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
