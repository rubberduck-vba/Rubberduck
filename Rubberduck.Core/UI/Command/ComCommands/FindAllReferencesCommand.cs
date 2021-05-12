using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    public class ProjectExplorerFindAllReferencesCommand : FindAllReferencesCommand
    {
        private readonly IParserStatusProvider _parserStatusProvider;
        private readonly ISelectedDeclarationProvider _finder;

        public ProjectExplorerFindAllReferencesCommand(
            IParserStatusProvider parserStatusProvider,
            IDeclarationFinderProvider declarationFinderProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IVBE vbe,
            FindAllReferencesAction finder,
            IVbeEvents vbeEvents)
            : base(parserStatusProvider, declarationFinderProvider, selectedDeclarationProvider, vbe, finder, vbeEvents)
        {
            _parserStatusProvider = parserStatusProvider;
            _finder = selectedDeclarationProvider;
        }

        protected override void OnExecute(object parameter)
        {
            if (_parserStatusProvider.Status != ParserState.Ready)
            {
                return;
            }

            var declaration = _finder.SelectedProjectExplorerModule();
            if (declaration == null)
            {
                return;
            }

            Service.FindAllReferences(declaration);
        }
    }

    /// <summary>
    /// A command that locates all references to a specified identifier, or of the active code module.
    /// </summary>
    public class FindAllReferencesCommand : ComCommandBase
    {
        private readonly IParserStatusProvider _parserStatusProvider;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IVBE _vbe;

        public FindAllReferencesCommand(
            IParserStatusProvider parserStatusProvider,
            IDeclarationFinderProvider declarationFinderProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IVBE vbe,
            FindAllReferencesAction finder, 
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _parserStatusProvider = parserStatusProvider;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _vbe = vbe;
            Service = finder;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        protected FindAllReferencesAction Service { get; }

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

            Service.FindAllReferences(declaration);
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            Declaration target = null;
            using (var activePane = _vbe.ActiveCodePane)
            using (var selectedComponent = _vbe.SelectedVBComponent)
            {
                if (activePane != null
                    && !activePane.IsWrappingNullReference
                    && (selectedComponent?.HasDesigner ?? false))
                {
                    using (var activeWindow = activePane.Window)
                    using (var designer = selectedComponent.DesignerWindow())
                    {
                        // Handle() is 0 for both windows, and IsVisible is true whenever the window is merely opened (active or not, regardless of state).
                        // Caption will be "UserForm1 (Code)" vs "UserForm1 (UserForm)"
                        if (designer.IsVisible && designer.Caption == activeWindow.Caption)
                        {
                            target = FindFormDesignerTarget(selectedComponent);
                        }
                    }
                }
            }

            return target ?? FindCodePaneTarget();
        }

        private Declaration FindCodePaneTarget() => _selectedDeclarationProvider.SelectedDeclaration();
        

        private Declaration FindFormDesignerTarget(IVBComponent component)
        {
            if (!component.HasDesigner)
            {
                return null;
            }

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
            // TODO determine if the above comment is still true.
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
