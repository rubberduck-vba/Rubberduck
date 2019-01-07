﻿using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringPresenterFactory<IReorderParametersPresenter> _factory;
        private ReorderParametersModel _model;
        private readonly IMessageBox _messageBox;
        private readonly IRewritingManager _rewritingManager;

        public ReorderParametersRefactoring(IVBE vbe, IRefactoringPresenterFactory<IReorderParametersPresenter> factory,
            IMessageBox messageBox, IRewritingManager rewritingManager)
        {
            _vbe = vbe;
            _factory = factory;
            _messageBox = messageBox;
            _rewritingManager = rewritingManager;
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

            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }

                var oldSelection = pane.GetQualifiedSelection();

                var rewriteSession = _rewritingManager.CheckOutCodePaneSession();
                AdjustReferences(_model.TargetDeclaration.References, rewriteSession);
                AdjustSignatures(rewriteSession);
                rewriteSession.TryRewrite();

                if (oldSelection.HasValue && !pane.IsWrappingNullReference)
                {
                    pane.Selection = oldSelection.Value.Selection;
                } 
            }
        }

        public void Refactor(QualifiedSelection target)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane == null || pane.IsWrappingNullReference)
                {
                    return;
                }

                pane.Selection = target.Selection;
            }
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            if (!ReorderParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                throw new ArgumentException("Invalid declaration type");
            }

            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane == null || pane.IsWrappingNullReference)
                {
                    return;
                }
                pane.Selection = target.QualifiedSelection.Selection;
            }
            Refactor();
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
                        _messageBox.NotifyWarn(RubberduckUI.ReorderPresenter_OptionalParametersMustBeLastError, RubberduckUI.ReorderParamsDialog_TitleText);
                        return false;
                    }
                }
            }

            var indexOfParamArray = _model.Parameters.FindIndex(param => param.IsParamArray);
            if (indexOfParamArray < 0 || indexOfParamArray == _model.Parameters.Count - 1)
            {
                return true;
            }

            _messageBox.NotifyWarn(RubberduckUI.ReorderPresenter_ParamArrayError, RubberduckUI.ReorderParamsDialog_TitleText);
            return false;
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references, IRewriteSession rewriteSession)
        {
            foreach (var reference in references.Where(item => item.Context != _model.TargetDeclaration.Context))
            {
                VBAParser.ArgumentListContext argumentList = null;
                var callStmt = reference.Context.GetAncestor<VBAParser.CallStmtContext>();
                if (callStmt != null)
                {
                    argumentList = CallStatement.GetArgumentList(callStmt);
                }
                
                if (argumentList == null)
                {
                    var indexExpression = reference.Context.GetAncestor<VBAParser.IndexExprContext>();
                    if (indexExpression != null)
                    {
                        argumentList = indexExpression.GetChild<VBAParser.ArgumentListContext>();
                    }
                }

                if (argumentList == null)
                {
                    continue; 
                }

                var module = reference.QualifiedModuleName;
                RewriteCall(argumentList, module, rewriteSession);
            }
        }

        private void RewriteCall(VBAParser.ArgumentListContext argList, QualifiedModuleName module, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(module);

            var args = argList.argument().Select((s, i) => new { Index = i, Text = s.GetText() }).ToList();
            for (var i = 0; i < _model.Parameters.Count; i++)
            {
                if (argList.argument().Length <= i)
                {
                    break;
                }

                var arg = argList.argument()[i];
                rewriter.Replace(arg, args.Single(s => s.Index == _model.Parameters[i].Index).Text);
            }
        }

        private void AdjustSignatures(IRewriteSession rewriteSession)
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
                    AdjustSignatures(setter, rewriteSession);
                    AdjustReferences(setter.References, rewriteSession);
                }

                var letter = _model.Declarations.FirstOrDefault(item => item.ParentScope == _model.TargetDeclaration.ParentScope &&
                              item.IdentifierName == _model.TargetDeclaration.IdentifierName &&
                              item.DeclarationType == DeclarationType.PropertyLet);

                if (letter != null)
                {
                    AdjustSignatures(letter, rewriteSession);
                    AdjustReferences(letter.References, rewriteSession);
                }
            }

            RewriteSignature(_model.TargetDeclaration, paramList, rewriteSession);

            foreach (var withEvents in _model.Declarations.Where(item => item.IsWithEvents && item.AsTypeName == _model.TargetDeclaration.ComponentName))
            {
                foreach (var reference in _model.Declarations.FindEventProcedures(withEvents))
                {
                    AdjustReferences(reference.References, rewriteSession);
                    AdjustSignatures(reference, rewriteSession);
                }
            }

            if (!(_model.TargetDeclaration is ModuleBodyElementDeclaration member) 
                || !(member.IsInterfaceImplementation || member.IsInterfaceMember))
            {
                return;
            }

            var implementations =
                _model.State.DeclarationFinder.FindInterfaceImplementationMembers(member.IsInterfaceMember
                    ? member
                    : member.InterfaceMemberImplemented);

            foreach (var interfaceImplentation in implementations)
            {
                AdjustReferences(interfaceImplentation.References, rewriteSession);
                AdjustSignatures(interfaceImplentation, rewriteSession);
            }
        }

        private void AdjustSignatures(Declaration declaration, IRewriteSession rewriteSession)
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

            RewriteSignature(declaration, paramList, rewriteSession);
        }

        private void RewriteSignature(Declaration target, VBAParser.ArgListContext paramList, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            var parameters = paramList.arg().Select((s, i) => new { Index = i, Text = s.GetText() }).ToList();
            for (var i = 0; i < _model.Parameters.Count; i++)
            {
                var param = paramList.arg()[i];
                rewriter.Replace(param, parameters.SingleOrDefault(s => s.Index == _model.Parameters[i].Index)?.Text);
            }
        }
    }
}
