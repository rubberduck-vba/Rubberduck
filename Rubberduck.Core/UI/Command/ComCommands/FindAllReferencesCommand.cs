using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    /// <summary>
    /// A command that locates all references to a specified identifier, or of the active code module.
    /// </summary>
    [ComVisible(false)]
    public class FindAllReferencesCommand : ComCommandBase
    {
        private readonly IParserStatusProvider _parserStatusProvider;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IVBE _vbe;
        private readonly FindAllReferencesService _finder;

        public FindAllReferencesCommand(
            IParserStatusProvider parserStatusProvider,
            IDeclarationFinderProvider declarationFinderProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IVBE vbe, 
            ISearchResultsWindowViewModel viewModel, 
            FindAllReferencesService finder, 
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _parserStatusProvider = parserStatusProvider;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _vbe = vbe;
            _finder = finder;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (_parserStatusProvider.Status != ParserState.Ready)
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
            if (_parserStatusProvider.Status != ParserState.Ready)
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
                using (var selectedComponent = _vbe.SelectedVBComponent)
                {
                    if (activePane != null
                        && !activePane.IsWrappingNullReference
                        && (selectedComponent?.HasDesigner ?? false))
                    {
                        return FindFormDesignerTarget(selectedComponent);
                    }
                }
            }

            return FindCodePaneTarget();
        }

        private Declaration FindCodePaneTarget()
        {
            return _selectedDeclarationProvider.SelectedDeclaration();
        }

        //Assumes the component has a designer.
        private Declaration FindFormDesignerTarget(IVBComponent component)
        {
            string projectId;
            using (var activeProject = _vbe.ActiveVBProject)
            {
                projectId = activeProject.ProjectId;
            }

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

            return _declarationFinderProvider.DeclarationFinder
                .MatchName(selectedName)
                .SingleOrDefault(m => m.ProjectId == projectId
                                      && m.DeclarationType.HasFlag(selectedType)
                                      && m.ComponentName == component.Name);
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
    }
}
