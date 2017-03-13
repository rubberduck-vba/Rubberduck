using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Common;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceFieldRefactoring : IRefactoring
    {
        private readonly IList<Declaration> _declarations;
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public IntroduceFieldRefactoring(IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
        {
            _declarations = state.AllDeclarations
                .Where(i => !i.IsBuiltIn && i.DeclarationType == DeclarationType.Variable)
                .ToList();

            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var selection = _vbe.ActiveCodePane.GetQualifiedSelection();

            if (!selection.HasValue)
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceField_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            Refactor(selection.Value);
        }

        public void Refactor(QualifiedSelection selection)
        {
            var target = _declarations.FindVariable(selection);

            if (target == null)
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            var rewriter = _state.GetRewriter(target);
            PromoteVariable(rewriter, target);

            _state.OnParseRequested(this);
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                throw new ArgumentException(@"Invalid declaration type", "target");
            }

            var rewriter = _state.GetRewriter(target);
            PromoteVariable(rewriter, target);
        }

        private void PromoteVariable(IModuleRewriter rewriter, Declaration target)
        {
            if (new[] { DeclarationType.ClassModule, DeclarationType.ProceduralModule }.Contains(target.ParentDeclaration.DeclarationType))
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            QualifiedSelection? oldSelection = null;
            if (_vbe.ActiveCodePane != null)
            {
                oldSelection = _vbe.ActiveCodePane.CodeModule.GetQualifiedSelection();
            }

            rewriter.Remove(target);
            AddField(rewriter, target);

            if (oldSelection.HasValue)
            {
                var module = oldSelection.Value.QualifiedName.Component.CodeModule;
                var pane = module.CodePane;
                {
                    pane.Selection = oldSelection.Value.Selection;
                }
            }
        }

        private void AddField(IModuleRewriter rewriter, Declaration target)
        {
            var content = "Private " + target.IdentifierName + " As " + target.AsTypeName + Environment.NewLine;
            rewriter.AppendToDeclarations(content);
        }
    }
}
