using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringPresenterFactory<IReorderParametersPresenter> _factory;
        private ReorderParametersModel _model;
        private readonly IMessageBox _messageBox;
        private readonly HashSet<IModuleRewriter> _rewriters = new HashSet<IModuleRewriter>();

        public ReorderParametersRefactoring(IVBE vbe, IRefactoringPresenterFactory<IReorderParametersPresenter> factory, IMessageBox messageBox)
        {
            _vbe = vbe;
            _factory = factory;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                return;
            }

            _model = presenter.Show();
            if (_model == null || !_model.Parameters.Where((param, index) => param.Index != index).Any() || !IsValidParamOrder())
            {
                return;
            }

            var pane = _vbe.ActiveCodePane;
            if (!pane.IsWrappingNullReference)
            {
                QualifiedSelection? oldSelection;
                var module = pane.CodeModule;
                {
                    oldSelection = module.GetQualifiedSelection();
                }

                AdjustReferences(_model.TargetDeclaration.References);
                AdjustSignatures();

                if (oldSelection.HasValue)
                {
                    pane.Selection = oldSelection.Value.Selection;
                }
            }

            foreach (var rewriter in _rewriters)
            {
                rewriter.Rewrite();
            }

            _model.State.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            var pane = _vbe.ActiveCodePane;
            {
                pane.Selection = target.Selection;
                Refactor();
            }
        }

        public void Refactor(Declaration target)
        {
            if (!ReorderParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                throw new ArgumentException("Invalid declaration type");
            }

            var pane = _vbe.ActiveCodePane;
            {
                pane.Selection = target.QualifiedSelection.Selection;
                Refactor();
            }
        }

        private bool IsValidParamOrder()
        {
            var indexOfFirstOptionalParam = _model.Parameters.FindIndex(param => param.IsOptional);
            if (indexOfFirstOptionalParam >= 0)
            {
                for (var index = indexOfFirstOptionalParam + 1; index < _model.Parameters.Count; index++)
                {
                    if (!_model.Parameters.ElementAt(index).IsOptional)
                    {
                        _messageBox.Show(RubberduckUI.ReorderPresenter_OptionalParametersMustBeLastError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return false;
                    }
                }
            }

            var indexOfParamArray = _model.Parameters.FindIndex(param => param.IsParamArray);
            if (indexOfParamArray >= 0 && indexOfParamArray != _model.Parameters.Count - 1)
            {
                _messageBox.Show(RubberduckUI.ReorderPresenter_ParamArrayError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            return true;
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references)
        {
            foreach (var reference in references.Where(item => item.Context != _model.TargetDeclaration.Context))
            {
                var module = reference.QualifiedModuleName.Component.CodeModule;
                VBAParser.ArgumentListContext argumentList = null;
                var callStmt = ParserRuleContextHelper.GetParent<VBAParser.CallStmtContext>(reference.Context);
                if (callStmt != null)
                {
                    argumentList = CallStatement.GetArgumentList(callStmt);
                }
                
                if (argumentList == null)
                {
                    var indexExpression = ParserRuleContextHelper.GetParent<VBAParser.IndexExprContext>(reference.Context);
                    if (indexExpression != null)
                    {
                        argumentList = ParserRuleContextHelper.GetChild<VBAParser.ArgumentListContext>(indexExpression);
                    }
                }

                if (argumentList == null) { continue; }
                RewriteCall(argumentList, module);
            }
        }

        private void RewriteCall(VBAParser.ArgumentListContext argList, ICodeModule module)
        {
            var rewriter = _model.State.GetRewriter(module.Parent);

            var args = argList.argument().Select((s, i) => new { Index = i, Text = s.GetText() }).ToList();
            for (var i = 0; i < _model.Parameters.Count; i++)
            {
                if (argList.argument().Count <= i)
                {
                    break;
                }

                var arg = argList.argument()[i];
                rewriter.Replace(arg, args.Single(s => s.Index == _model.Parameters[i].Index).Text);
            }

            _rewriters.Add(rewriter);
        }

        private void AdjustSignatures()
        {
            var proc = (dynamic)_model.TargetDeclaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();

            // if we are reordering a property getter, check if we need to reorder a letter/setter too
            if (_model.TargetDeclaration.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = _model.Declarations.FirstOrDefault(item => item.ParentScope == _model.TargetDeclaration.ParentScope &&
                                              item.IdentifierName == _model.TargetDeclaration.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertySet);

                if (setter != null)
                {
                    AdjustSignatures(setter);
                    AdjustReferences(setter.References);
                }

                var letter = _model.Declarations.FirstOrDefault(item => item.ParentScope == _model.TargetDeclaration.ParentScope &&
                              item.IdentifierName == _model.TargetDeclaration.IdentifierName &&
                              item.DeclarationType == DeclarationType.PropertyLet);

                if (letter != null)
                {
                    AdjustSignatures(letter);
                    AdjustReferences(letter.References);
                }
            }

            RewriteSignature(_model.TargetDeclaration, paramList);

            foreach (var withEvents in _model.Declarations.Where(item => item.IsWithEvents && item.AsTypeName == _model.TargetDeclaration.ComponentName))
            {
                foreach (var reference in _model.Declarations.FindEventProcedures(withEvents))
                {
                    AdjustReferences(reference.References);
                    AdjustSignatures(reference);
                }
            }

            var interfaceImplementations = _model.Declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.ProjectId == _model.TargetDeclaration.ProjectId &&
                                                               item.IdentifierName == _model.TargetDeclaration.ComponentName + "_" + _model.TargetDeclaration.IdentifierName);
            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustReferences(interfaceImplentation.References);
                AdjustSignatures(interfaceImplentation);
            }
        }

        private void AdjustSignatures(Declaration declaration)
        {
            var proc = (dynamic) declaration.Context.Parent;
            VBAParser.ArgListContext paramList;

            if (declaration.DeclarationType == DeclarationType.PropertySet ||
                declaration.DeclarationType == DeclarationType.PropertyLet)
            {
                paramList = (VBAParser.ArgListContext) proc.children[0].argList();
            }
            else
            {
                paramList = (VBAParser.ArgListContext) proc.subStmt().argList();
            }

            RewriteSignature(declaration, paramList);
        }

        private void RewriteSignature(Declaration target, VBAParser.ArgListContext paramList)
        {
            var rewriter = _model.State.GetRewriter(target);

            var parameters = paramList.arg().Select((s, i) => new { Index = i, Text = s.GetText() }).ToList();
            for (var i = 0; i < _model.Parameters.Count; i++)
            {
                var param = paramList.arg()[i];
                rewriter.Replace(param, parameters.SingleOrDefault(s => s.Index == _model.Parameters[i].Index)?.Text);
            }

            _rewriters.Add(rewriter);
        }
    }
}
