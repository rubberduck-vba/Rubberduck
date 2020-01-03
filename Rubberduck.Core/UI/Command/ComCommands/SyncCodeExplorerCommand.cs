using System.Linq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.ComCommands
{
    public interface ICodeExplorerSyncProvider
    {
        SyncCodeExplorerCommand GetSyncCommand(CodeExplorerViewModel explorer);
    }

    public class CodeExplorerSyncProvider : ICodeExplorerSyncProvider
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IVbeEvents _vbeEvents;

        public CodeExplorerSyncProvider(
            IVBE vbe, 
            RubberduckParserState state, 
            IVbeEvents vbeEvents)
        {
            _vbe = vbe;
            _state = state;
            _vbeEvents = vbeEvents;
        }

        public SyncCodeExplorerCommand GetSyncCommand(CodeExplorerViewModel explorer)
        {
            var selectionService = new SelectionService(_vbe, _state.ProjectsProvider);
            var selectedDeclarationService = new SelectedDeclarationProvider(selectionService, _state);
            return new SyncCodeExplorerCommand(_vbe, _state, _state, selectedDeclarationService, explorer, _vbeEvents);
        }
    }

    public class SyncCodeExplorerCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IParserStatusProvider _parserStatusProvider;
        private readonly CodeExplorerViewModel _explorer;

        public SyncCodeExplorerCommand(
            IVBE vbe,
            IDeclarationFinderProvider declarationFinderProvider, 
            IParserStatusProvider parserStatusProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            CodeExplorerViewModel explorer, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _parserStatusProvider = parserStatusProvider;
            _explorer = explorer;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _parserStatusProvider.Status == ParserState.Ready 
                   && !_explorer.IsBusy 
                   && FindTargetNode() != null;
        }

        protected override void OnExecute(object parameter)
        {
            var target = FindTargetNode();

            if (target == null)
            {
                return;
            }

            _explorer.SelectedItem = target;
        }

        private ICodeExplorerNode FindTargetNode()
        {
            var targetDeclaration = FindTargetDeclaration();
            return targetDeclaration != null
                ? _explorer.FindVisibleNodeForDeclaration(targetDeclaration)
                : null;
        }

        private Declaration FindTargetDeclaration()
        {
            return _selectedDeclarationProvider.SelectedDeclaration()
                ?? ActiveProjectDeclaration();
        }

        private Declaration ActiveProjectDeclaration()
        {
            var projectId = ActiveProjectId();

            if (projectId == null)
            {
                return null;
            }

            return _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Project)
                .FirstOrDefault(item => item.ProjectId.Equals(projectId));
        }

        private string ActiveProjectId()
        {
            using (var project = _vbe.ActiveVBProject)
            {
                if (project == null || project.IsWrappingNullReference)
                {
                    return null;
                }

                return project.ProjectId;
            }
        }
    }
}
