using System.Linq;
using System.Windows;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    public interface ICodeExplorerSyncProvider
    {
        SyncCodeExplorerCommand GetSyncCommand(CodeExplorerViewModel explorer);
    }

    public class CodeExplorerSyncProvider : ICodeExplorerSyncProvider
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;

        public CodeExplorerSyncProvider(IVBE vbe, RubberduckParserState state)
        {
            _vbe = vbe;
            _state = state;
        }

        public SyncCodeExplorerCommand GetSyncCommand(CodeExplorerViewModel explorer)
        {
            return new SyncCodeExplorerCommand(_vbe, _state, explorer);
        }
    }

    public class SyncCodeExplorerCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly CodeExplorerViewModel _explorer;

        public SyncCodeExplorerCommand(IVBE vbe, RubberduckParserState state, CodeExplorerViewModel explorer) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
            _explorer = explorer;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready || 
                _explorer.IsBusy || 
                FindTargetNode() == null)
            {
                return false;
            }

            return true;
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
            using (var active = _vbe.ActiveCodePane)
            {
                if (active == null || active.IsWrappingNullReference)
                {
                    using (var project = _vbe.ActiveVBProject)
                    {
                        if (project == null || project.IsWrappingNullReference)
                        {
                            return null;
                        }

                        var declaration = _state.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                            .FirstOrDefault(item => item.ProjectId.Equals(project.ProjectId));

                        return _explorer.FindVisibleNodeForDeclaration(declaration);
                    }
                }

                var selected = _state.DeclarationFinder?.FindSelectedDeclaration(active);

                return selected == null ? null : _explorer.FindVisibleNodeForDeclaration(selected);
            }
        } 
    }
}
