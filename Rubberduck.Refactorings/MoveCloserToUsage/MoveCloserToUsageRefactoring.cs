using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.Resources;
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
            var qualifiedSelection = _vbe.GetActiveSelection();

            if (qualifiedSelection != null)
            {
                Refactor(_declarations.FindVariable(qualifiedSelection.Value));
            }
            else
            {
                _messageBox.NotifyWarn(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.MoveCloserToUsage_Caption);
            }
        }

        public void Refactor(QualifiedSelection selection)
        {
            _target = _declarations.FindVariable(selection);

            if (_target == null)
            {
                _messageBox.NotifyWarn(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            MoveCloserToUsage();
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                _messageBox.NotifyWarn(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
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
                _messageBox.NotifyWarn(message, RubberduckUI.MoveCloserToUsage_Caption);
                return;
            }

            if (TargetIsReferencedFromMultipleMethods(_target))
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetIsUsedInMultipleMethods,
                    _target.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.MoveCloserToUsage_Caption);

                return;
            }

            QualifiedSelection? oldSelection = null;
            using (var pane = _vbe.ActiveCodePane)
            {
                if (!pane.IsWrappingNullReference)
                {
                    oldSelection = pane.GetQualifiedSelection();
                }

                InsertNewDeclaration();
                RemoveOldDeclaration();
                UpdateOtherModules();

                if (oldSelection.HasValue && !pane.IsWrappingNullReference)
                {
                    pane.Selection = oldSelection.Value.Selection;
                }
            }
            foreach (var rewriter in _rewriters)
            {
                rewriter.Rewrite();
            }
            Reparse();
        }

        private void UpdateOtherModules()
        {
            QualifiedSelection? oldSelection = null;
            using (var pane = _vbe.ActiveCodePane)
            {
                if (!pane.IsWrappingNullReference)
                {
                    oldSelection = pane.GetQualifiedSelection();
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
        }

        private void InsertNewDeclaration()
        {
            var subscripts = _target.Context.GetDescendent<VBAParser.SubscriptsContext>()?.GetText() ?? string.Empty;
            var identifier = _target.IsArray ? $"{_target.IdentifierName}({subscripts})" : _target.IdentifierName;

            var newVariable = _target.AsTypeContext is null
                ? $"{Tokens.Dim} {identifier} {Tokens.As} {Tokens.Variant}{Environment.NewLine}"
                : $"{Tokens.Dim} {identifier} {Tokens.As} {(_target.IsSelfAssigned ? Tokens.New + " " : string.Empty)}{_target.AsTypeNameWithoutArrayDesignator}{Environment.NewLine}";

            var firstReference = _target.References.OrderBy(r => r.Selection.StartLine).First();

            RuleContext expression = firstReference.Context;
            while (!(expression is VBAParser.BlockStmtContext))
            {
                expression = expression.Parent;
            }

            var insertionIndex = (expression as ParserRuleContext).Start.TokenIndex;
            int indentLength;
            using (var pane = _vbe.ActiveCodePane)
            {
                using (var codeModule = pane.CodeModule)
                {
                    var firstReferenceLine = codeModule.GetLines((expression as ParserRuleContext).Start.Line, 1);
                    indentLength = firstReferenceLine.Length - firstReferenceLine.TrimStart().Length;
                }
            }
            var padding = new string(' ', indentLength);

            var rewriter = _state.GetRewriter(firstReference.QualifiedModuleName);
            rewriter.InsertBefore(insertionIndex, newVariable + padding);

            _rewriters.Add(rewriter);
        }

        private void RemoveOldDeclaration()
        {
            var rewriter = _state.GetRewriter(_target);
            rewriter.Remove(_target);

            _rewriters.Add(rewriter);
        }

        private void UpdateCallsToOtherModule(IEnumerable<IdentifierReference> references)
        {
            foreach (var reference in references.OrderByDescending(o => o.Selection.StartLine).ThenByDescending(t => t.Selection.StartColumn))
            {
                // todo: Grab `GetAncestor` and use that
                var parent = reference.Context.Parent;
                while (!(parent is VBAParser.MemberAccessExprContext) && parent.Parent != null)
                {
                    parent = parent.Parent;
                }

                if (!(parent is VBAParser.MemberAccessExprContext))
                {
                    continue;
                }

                // member access might be to something unrelated to the rewritten target.
                // check we're not accidentally overwriting some other member-access who just happens to be a parent context
                var memberAccessContext = (VBAParser.MemberAccessExprContext)parent;
                if (memberAccessContext.unrestrictedIdentifier().GetText() != _target.IdentifierName)
                {
                    continue;
                }

                var rewriter = _state.GetRewriter(reference.QualifiedModuleName);
                var tokenInterval = Interval.Of(parent.SourceInterval.a, reference.Context.SourceInterval.b);
                rewriter.Replace(tokenInterval, reference.IdentifierName);

                _rewriters.Add(rewriter);
            }
        }

        private void Reparse()
        {
            _state.OnParseRequested(this);
        }
    }
}
