﻿using System.Linq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
            return new SyncCodeExplorerCommand(_vbe, _state, explorer, _vbeEvents);
        }
    }

    public class SyncCodeExplorerCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly CodeExplorerViewModel _explorer;

        public SyncCodeExplorerCommand(
            IVBE vbe, 
            RubberduckParserState state, 
            CodeExplorerViewModel explorer, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _state = state;
            _explorer = explorer;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready 
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
