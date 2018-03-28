using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that finds all implementations of a specified method, or of the active interface module.
    /// </summary>
    [ComVisible(false)]
    public class FindAllImplementationsCommand : CommandBase, IDisposable
    {
        private readonly INavigateCommand _navigateCommand;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;
        private readonly IVBE _vbe;
        private readonly IUiDispatcher _uiDispatcher;

        private new static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public FindAllImplementationsCommand(INavigateCommand navigateCommand, IMessageBox messageBox,
            RubberduckParserState state, IVBE vbe, ISearchResultsWindowViewModel viewModel,
            SearchResultPresenterInstanceManager presenterService, IUiDispatcher uiDispatcher)
             : base(LogManager.GetCurrentClassLogger())
        {
            _navigateCommand = navigateCommand;
            _messageBox = messageBox;
            _state = state;
            _vbe = vbe;
            _viewModel = viewModel;
            _presenterService = presenterService;
            _uiDispatcher = uiDispatcher;

            _state.StateChanged += _state_StateChanged;
        }

        private Declaration FindNewDeclaration(Declaration declaration)
        {
            return _state.AllUserDeclarations.SingleOrDefault(item =>
                        item.ProjectId == declaration.ProjectId &&
                        item.ComponentName == declaration.ComponentName &&
                        item.ParentScope == declaration.ParentScope &&
                        item.IdentifierName == declaration.IdentifierName &&
                        item.DeclarationType == declaration.DeclarationType);
        }

        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready) { return; }

            if (_viewModel == null) { return; }

            _uiDispatcher.InvokeAsync(UpdateTab);
        }

        private void UpdateTab()
        {
            try
            {
                var findImplementationsTabs = _viewModel.Tabs.Where(
                    t => t.Header.StartsWith(RubberduckUI.AllImplementations_Caption.Replace("'{0}'", ""))).ToList();

                foreach (var tab in findImplementationsTabs)
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
                Logger.Error(exception, "Exception thrown while trying to update the find implementations tab.");
            }
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            using (var codePane = _vbe.ActiveCodePane)
            {
                if (codePane == null || codePane.IsWrappingNullReference || _state.Status != ParserState.Ready)
                {
                    return false;
                }

                var target = FindTarget(parameter);
                var canExecute = target != null;

                return canExecute;
            }
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
                Console.WriteLine(e);
            }
        }

        private SearchResultsViewModel CreateViewModel(Declaration target)
        {
            var results = FindImplementations(target).Select(declaration =>
                new SearchResultItem(
                    declaration.ParentScopeDeclaration,
                    new NavigateCodeEventArgs(declaration.QualifiedName.QualifiedModuleName, declaration.Selection),
                    GetModuleLine(declaration.QualifiedName.QualifiedModuleName, declaration.Selection.StartLine)));

            var viewModel = new SearchResultsViewModel(_navigateCommand,
                string.Format(RubberduckUI.SearchResults_AllImplementationsTabFormat, target.IdentifierName), target, results);

            return viewModel;
        }

        private string GetModuleLine(QualifiedModuleName module, int line)
        {
            using (var codeModule = _state.ProjectsProvider.Component(module).CodeModule)
            {
                return codeModule.GetLines(line, 1).Trim();
            }
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            using (var activePane = _vbe.ActiveCodePane)
            {
                return _state.FindSelectedDeclaration(activePane);
            }
        }

        private IEnumerable<Declaration> FindImplementations(Declaration target)
        {
            var items = _state.AllDeclarations;
            var implementations = (target.DeclarationType == DeclarationType.ClassModule
                ? FindAllImplementationsOfClass(target, items, out _)
                : FindAllImplementationsOfMember(target, items, out _)) ?? new List<Declaration>();

            return implementations;
        }

        private IEnumerable<Declaration> FindAllImplementationsOfClass(Declaration target, IEnumerable<Declaration> declarations, out string name)
        {
            if (target.DeclarationType != DeclarationType.ClassModule)
            {
                name = string.Empty;
                return Enumerable.Empty<Declaration>();
            }

            var identifiers = declarations as IList<Declaration> ?? declarations.ToList();

            var result = target.References
                .Where(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext)
                .SelectMany(reference => identifiers.Where(identifier => identifier.IdentifierName == reference.QualifiedModuleName.ComponentName))
                .ToList();

            name = target.ComponentName;
            return result;
        }

        private IEnumerable<Declaration> FindAllImplementationsOfMember(Declaration target, IEnumerable<Declaration> declarations, out string name)
        {
            if (!target.DeclarationType.HasFlag(DeclarationType.Member))
            {
                name = string.Empty;
                return Enumerable.Empty<Declaration>();
            }

            var items = declarations as IList<Declaration> ?? declarations.ToList();

            var isInterface = items.FindInterfaces()
                .Select(i => i.QualifiedName.QualifiedModuleName.ToString())
                .Contains(target.QualifiedName.QualifiedModuleName.ToString());

            if (isInterface)
            {
                name = target.ComponentName + "." + target.IdentifierName;
                return items.FindInterfaceImplementationMembers(target.IdentifierName)
                       .Where(item => item.IdentifierName == target.ComponentName + "_" + target.IdentifierName);
            }

            var member = items.FindInterfaceMember(target);
            if (member == null)
            {
                name = string.Empty;
                return Enumerable.Empty<Declaration>();
            }
            name = member.ComponentName + "." + member.IdentifierName;
            return items.FindInterfaceImplementationMembers(member.IdentifierName)
                   .Where(item => item.IdentifierName == member.ComponentName + "_" + member.IdentifierName);
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
