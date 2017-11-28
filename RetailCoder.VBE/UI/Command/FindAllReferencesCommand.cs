using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that locates all references to a specified identifier, or of the active code module.
    /// </summary>
    [ComVisible(false)]
    public class FindAllReferencesCommand : CommandBase, IDisposable
    {
        private readonly INavigateCommand _navigateCommand;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;
        private readonly IVBE _vbe;

        public FindAllReferencesCommand(INavigateCommand navigateCommand, IMessageBox messageBox,
            RubberduckParserState state, IVBE vbe, ISearchResultsWindowViewModel viewModel,
            SearchResultPresenterInstanceManager presenterService)
             : base(LogManager.GetCurrentClassLogger())
        {
            _navigateCommand = navigateCommand;
            _messageBox = messageBox;
            _state = state;
            _vbe = vbe;
            _viewModel = viewModel;
            _presenterService = presenterService;

            _state.StateChanged += _state_StateChanged;
        }

        private Declaration FindNewDeclaration(Declaration declaration)
        {
            return _state.DeclarationFinder
                .MatchName(declaration.IdentifierName)
                .SingleOrDefault(d => d.ProjectId == declaration.ProjectId
                    && d.ComponentName == declaration.ComponentName
                    && d.ParentScope == declaration.ParentScope
                    && d.DeclarationType == declaration.DeclarationType);            
        }

        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready) { return; }

            if (_viewModel == null) { return; }

            UiDispatcher.InvokeAsync(UpdateTab);
        }

        private void UpdateTab()
        {
            try
            {
                var findReferenceTabs = _viewModel.Tabs.Where(
                    t => t.Header.StartsWith(RubberduckUI.AllReferences_Caption.Replace("'{0}'", ""))).ToList();

                foreach (var tab in findReferenceTabs)
                {
                    var newTarget = FindNewDeclaration(tab.Target);
                    if (newTarget == null)
                    {
                        tab.CloseCommand.Execute(null);
                        return;
                    }

                    var vm = CreateViewModel(newTarget);
                    if (vm.SearchResults.Any())
                    {
                        tab.SearchResults = vm.SearchResults;
                        tab.Target = vm.Target;
                    }
                    else
                    {
                        tab.CloseCommand.Execute(null);
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Exception thrown while trying to update the find all references tab.");
            }
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready ||
                (_vbe.ActiveCodePane == null && !(_vbe.SelectedVBComponent?.HasDesigner ?? false)))
            {
                return false;
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

            var viewModel = CreateViewModel(declaration);
            if (!viewModel.SearchResults.Any())
            {
                _messageBox.Show(string.Format(RubberduckUI.AllReferences_NoneFound, declaration.IdentifierName), RubberduckUI.Rubberduck, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (viewModel.SearchResults.Count == 1)
            {
                _navigateCommand.Execute(viewModel.SearchResults.Single().GetNavigationArgs());
                return;
            }

            _viewModel.AddTab(viewModel);
            _viewModel.SelectedTab = viewModel;

            try
            {
                var presenter = _presenterService.Presenter(_viewModel);
                presenter.Show();
            }
            catch (Exception e)
            {
                Logger.Error(e);
            }
        }

        private SearchResultsViewModel CreateViewModel(Declaration declaration)
        {
            var results = declaration.References.Distinct().Select(reference =>
                new SearchResultItem(
                    reference.ParentNonScoping,
                    new NavigateCodeEventArgs(reference.QualifiedModuleName, reference.Selection), 
                    reference.QualifiedModuleName.Component.CodeModule.GetLines(reference.Selection.StartLine, 1).Trim()));
            
            var viewModel = new SearchResultsViewModel(_navigateCommand,
                string.Format(RubberduckUI.SearchResults_AllReferencesTabFormat, declaration.IdentifierName), declaration, results);

            return viewModel;
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            return _vbe.ActiveCodePane != null && (_vbe.SelectedVBComponent?.HasDesigner ?? false)
                ? FindFormDesignerTarget()
                : FindCodePaneTarget();
        }

        private Declaration FindCodePaneTarget()
        {
            return _state.FindSelectedDeclaration(_vbe.ActiveCodePane);
        }

        private Declaration FindFormDesignerTarget(QualifiedModuleName? qualifiedModuleName = null)
        {            
            (var projectId, var component) = qualifiedModuleName.HasValue
                ? (qualifiedModuleName.Value.ProjectId, qualifiedModuleName.Value.Component)
                : (_vbe.ActiveVBProject.ProjectId, _vbe.SelectedVBComponent);

            if (component?.HasDesigner ?? false)
            {
                if (qualifiedModuleName.HasValue)
                {
                    return _state.DeclarationFinder
                        .MatchName(qualifiedModuleName.Value.Name)
                        .SingleOrDefault(m => m.ProjectId == projectId
                            && m.DeclarationType.HasFlag(qualifiedModuleName.Value.ComponentType)
                            && m.ComponentName == component.Name);
                }

                var selectedCount = component.SelectedControls.Count;                
                if (selectedCount > 1) { return null; }

                // Cannot use DeclarationType.UserForm, parser only assigns UserForms the ClassModule flag
                (var selectedType, var selectedName) = selectedCount == 0
                    ? (DeclarationType.ClassModule, component.Name)
                    : (DeclarationType.Control, component.SelectedControls[0].Name);
                
                return _state.DeclarationFinder
                    .MatchName(selectedName)
                    .SingleOrDefault(m => m.ProjectId == projectId
                        && m.DeclarationType.HasFlag(selectedType)
                        && m.ComponentName == component.Name);                
            }
            return null;
        }


        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= _state_StateChanged;
            }
        }
    }
}
