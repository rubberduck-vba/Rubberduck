﻿using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorMoveCloserToUsageCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _msgbox;

        public RefactorMoveCloserToUsageCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgbox)
            :base(vbe)
        {
            _state = state;
            _msgbox = msgbox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            using (var activePane = Vbe.ActiveCodePane)
            {
                if (activePane == null || activePane .IsWrappingNullReference || _state.Status != ParserState.Ready)
                {
                    return false;
                }

                var target = _state.FindSelectedDeclaration(activePane);
                return target != null
                       && !_state.IsNewOrModified(target.QualifiedModuleName)
                       && (target.DeclarationType == DeclarationType.Variable ||
                           target.DeclarationType == DeclarationType.Constant)
                       && target.References.Any();
            }
        }

        protected override void OnExecute(object parameter)
        {
            var selection = Vbe.GetActiveSelection();

            if (selection.HasValue)
            {
                var refactoring = new MoveCloserToUsageRefactoring(Vbe, _state, _msgbox);
                refactoring.Refactor(selection.Value);
            }
        }
    }
}
