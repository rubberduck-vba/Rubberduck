using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageRefactoring : IRefactoring
    {
        private readonly List<Declaration> _declarations;
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private Declaration _target;
        
        private readonly HashSet<IModuleRewriter> _rewriters = new HashSet<IModuleRewriter>();

        public MoveCloserToUsageRefactoring(IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
        {
            _declarations = state.AllUserDeclarations.ToList();
            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var qualifiedSelection = _vbe.ActiveCodePane.CodeModule.GetQualifiedSelection();
            if (qualifiedSelection != null)
            {
                Refactor(_declarations.FindVariable(qualifiedSelection.Value));
            }
            else
            {
                _messageBox.Show(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.MoveCloserToUsage_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }
        }

        public void Refactor(QualifiedSelection selection)
        {
            _target = _declarations.FindVariable(selection);

            if (_target == null)
            {
                _messageBox.Show(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            MoveCloserToUsage();
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                _messageBox.Show(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            _target = target;
            MoveCloserToUsage();
        }

        private bool TargetIsReferencedFromMultipleMethods(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();

            return firstReference != null && target.References.Any(r => !Equals(r.ParentScoping, firstReference.ParentScoping));
        }

        private void MoveCloserToUsage()
        {
            if (!_target.References.Any())
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetHasNoReferences, _target.IdentifierName);

                _messageBox.Show(message, RubberduckUI.MoveCloserToUsage_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);

                return;
            }

            if (TargetIsReferencedFromMultipleMethods(_target))
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetIsUsedInMultipleMethods,
                    _target.IdentifierName);
                _messageBox.Show(message, RubberduckUI.MoveCloserToUsage_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);

                return;
            }

            QualifiedSelection? oldSelection = null;
            var pane = _vbe.ActiveCodePane;
            var module = pane.CodeModule;
            if (!module.IsWrappingNullReference)
            {
                oldSelection = module.GetQualifiedSelection();
            }
            
            InsertNewDeclaration();
            RemoveOldDeclaration();
            UpdateOtherModules();

            if (oldSelection.HasValue)
            {
                pane.Selection = oldSelection.Value.Selection;
            }

            foreach (var rewriter in _rewriters)
            {
                rewriter.Rewrite();
            }
        }

        private void UpdateOtherModules()
        {
            QualifiedSelection? oldSelection = null;
            var pane = _vbe.ActiveCodePane;
            var module = pane.CodeModule;
            if (!module.IsWrappingNullReference)
            {
                oldSelection = module.GetQualifiedSelection();
            }

            var newTarget = _state.DeclarationFinder.MatchName(_target.IdentifierName).FirstOrDefault(
                item => item.ComponentName == _target.ComponentName &&
                        item.ParentScope == _target.ParentScope &&
                        item.ProjectId == _target.ProjectId &&
                        Equals(item.Selection, _target.Selection));

            if (newTarget != null)
            {
                UpdateCallsToOtherModule(newTarget.References.ToList());
            }

            if (oldSelection.HasValue)
            {
                pane.Selection = oldSelection.Value.Selection;
            }
        }

        private void InsertNewDeclaration()
        {
            var newVariable = $"Dim {_target.IdentifierName} As {_target.AsTypeName}{Environment.NewLine}";

            var firstReference = _target.References.OrderBy(r => r.Selection.StartLine).First();

            RuleContext expression = firstReference.Context;
            while (!(expression is VBAParser.BlockStmtContext))
            {
                expression = expression.Parent;
            }

            var insertionIndex = (expression as ParserRuleContext).Start.TokenIndex;

            var rewriter = _state.GetRewriter(firstReference.QualifiedModuleName);
            rewriter.InsertBefore(insertionIndex, newVariable);

            _rewriters.Add(rewriter);
        }

        private void RemoveOldDeclaration()
        {
            var rewriter = _state.GetRewriter(_target);
            rewriter.Remove(_target);

            _rewriters.Add(rewriter);
        }

        private void UpdateCallsToOtherModule(List<IdentifierReference> references)
        {
            foreach (var reference in references.OrderByDescending(o => o.Selection.StartLine).ThenByDescending(t => t.Selection.StartColumn))
            {
                var parent = reference.Context.Parent;
                while (!(parent is VBAParser.MemberAccessExprContext) && parent.Parent != null)
                {
                    parent = parent.Parent;
                }

                if (!(parent is VBAParser.MemberAccessExprContext))
                {
                    continue;
                }

                var rewriter = _state.GetRewriter(reference.QualifiedModuleName);
                rewriter.Replace(parent as ParserRuleContext, reference.IdentifierName);

                _rewriters.Add(rewriter);
            }
        }
    }
}
