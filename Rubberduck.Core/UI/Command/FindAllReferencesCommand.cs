using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that locates all references to a specified identifier, or of the active code module.
    /// </summary>
    [ComVisible(false)]
    public class FindAllReferencesCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;
        private readonly FindAllReferencesService _finder;

        public FindAllReferencesCommand(RubberduckParserState state, IVBE vbe, ISearchResultsWindowViewModel viewModel, FindAllReferencesService finder)
            : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
            _vbe = vbe;
            _finder = finder;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            using (var activePane = _vbe.ActiveCodePane)
            {
                using (var selectedComponent = _vbe.SelectedVBComponent)
                {
                    if ((activePane == null || activePane.IsWrappingNullReference)
                        && !(selectedComponent?.HasDesigner ?? false))
                    {
                        return false;
                    }
                }
            }

            var target = FindTarget(parameter);
            var canExecute = target != null;

            return canExecute;
        }

        protected override void OnExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            var declaration = FindTarget(parameter);
            if (declaration == null)
            {
                return;
            }

            _finder.FindAllReferences(declaration);
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            using (var activePane = _vbe.ActiveCodePane)
            {
                bool findDesigner;
                using (var selectedComponent = _vbe.SelectedVBComponent)
                {
                    findDesigner = activePane != null && !activePane.IsWrappingNullReference
                                                      && (selectedComponent?.HasDesigner ?? false);
                }

                return findDesigner
                    ? FindFormDesignerTarget()
                    : FindCodePaneTarget(activePane);
            }
        }

        private Declaration FindCodePaneTarget(ICodePane codePane)
        {
            return _state.FindSelectedDeclaration(codePane);
        }

        private Declaration FindFormDesignerTarget(QualifiedModuleName? qualifiedModuleName = null)
        {
            if (qualifiedModuleName.HasValue)
            {
                return FindFormDesignerTarget(qualifiedModuleName.Value);
            }

            string projectId;
            using (var activeProject = _vbe.ActiveVBProject)
            {
                projectId = activeProject.ProjectId;
            }

            using (var component = _vbe.SelectedVBComponent)
            {
                if (component?.HasDesigner ?? false)
                {
                    DeclarationType selectedType;
                    string selectedName;
                    using (var selectedControls = component.SelectedControls)
                    {
                        var selectedCount = selectedControls.Count;
                        if (selectedCount > 1)
                        {
                            return null;
                        }

                        (selectedType, selectedName) = GetSelectedName(component, selectedControls, selectedCount);
                    }

                    return _state.DeclarationFinder
                        .MatchName(selectedName)
                        .SingleOrDefault(m => m.ProjectId == projectId
                                              && m.DeclarationType.HasFlag(selectedType)
                                              && m.ComponentName == component.Name);
                }

                return null;
            }
        }

        private static (DeclarationType, string Name) GetSelectedName(IVBComponent component, IControls selectedControls, int selectedCount)
        {
            // Cannot use DeclarationType.UserForm, parser only assigns UserForms the ClassModule flag
            if (selectedCount == 0)
            {
                return (DeclarationType.ClassModule, component.Name);
            }

            using (var firstSelectedControl = selectedControls[0])
            {
                return (DeclarationType.Control, firstSelectedControl.Name);
            }
        }

        private Declaration FindFormDesignerTarget(QualifiedModuleName qualifiedModuleName)
        {
            var projectId = qualifiedModuleName.ProjectId;
            var component = _state.ProjectsProvider.Component(qualifiedModuleName);

            if (component?.HasDesigner ?? false)
            {
                return _state.DeclarationFinder
                    .MatchName(qualifiedModuleName.Name)
                    .SingleOrDefault(m => m.ProjectId == projectId
                                          && m.DeclarationType.HasFlag(qualifiedModuleName.ComponentType)
                                          && m.ComponentName == component.Name);
            }

            return null;
        }
    }
}
