﻿using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Rewriter;
using System.Collections.Generic;
using System.Linq;
 using Rubberduck.Parsing.VBA;
 using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersRefactoring : InteractiveRefactoringBase<IReorderParametersPresenter, ReorderParametersModel>
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public ReorderParametersRefactoring(RubberduckParserState state, IRefactoringPresenterFactory factory, IMessageBox messageBox, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _state = state;
            _messageBox = messageBox;
        }

        public override void Refactor(QualifiedSelection target)
        {
            Model = InitializeModel(target);
            if (Model == null)
            {
                return;
            }

            using (var container = PresenterFactory(Model))
            {
                var presenter = container.Value;
                if (presenter == null)
                {
                    return;
                }

                Model = presenter.Show();
                if (Model == null)
                {
                    return;
                }

                RefactorImpl(presenter);
            }
        }

        private ReorderParametersModel InitializeModel(QualifiedSelection targetSelection)
        {
            return new ReorderParametersModel(_state, targetSelection);
        }

        protected override void RefactorImpl(IReorderParametersPresenter presenter)
        {
            if (!Model.Parameters.Where((param, index) => param.Index != index).Any()
                || !IsValidParamOrder())
            {
                return;
            }

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            AdjustReferences(Model.TargetDeclaration.References, rewriteSession);
            AdjustSignatures(rewriteSession);
            rewriteSession.TryRewrite();
        }

        protected override ReorderParametersModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                return null;
            }

            if (!ReorderParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                return null;
            }

            return InitializeModel(target.QualifiedSelection);
        }

        private bool IsValidParamOrder()
        {
            var indexOfFirstOptionalParam = Model.Parameters.FindIndex(param => param.IsOptional);
            if (indexOfFirstOptionalParam >= 0)
            {
                for (var index = indexOfFirstOptionalParam + 1; index < Model.Parameters.Count; index++)
                {
                    if (!Model.Parameters.ElementAt(index).IsOptional)
                    {
                        _messageBox.NotifyWarn(RubberduckUI.ReorderPresenter_OptionalParametersMustBeLastError, RubberduckUI.ReorderParamsDialog_TitleText);
                        return false;
                    }
                }
            }

            var indexOfParamArray = Model.Parameters.FindIndex(param => param.IsParamArray);
            if (indexOfParamArray < 0 || indexOfParamArray == Model.Parameters.Count - 1)
            {
                return true;
            }

            _messageBox.NotifyWarn(RubberduckUI.ReorderPresenter_ParamArrayError, RubberduckUI.ReorderParamsDialog_TitleText);
            return false;
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references, IRewriteSession rewriteSession)
        {
            foreach (var reference in references.Where(item => item.Context != Model.TargetDeclaration.Context))
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
            for (var i = 0; i < Model.Parameters.Count; i++)
            {
                if (argList.argument().Length <= i)
                {
                    break;
                }

                var arg = argList.argument()[i];
                rewriter.Replace(arg, args.Single(s => s.Index == Model.Parameters[i].Index).Text);
            }
        }

        private void AdjustSignatures(IRewriteSession rewriteSession)
        {
            var proc = (dynamic)Model.TargetDeclaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();

            // if we are reordering a property getter, check if we need to reorder a letter/setter too
            if (Model.TargetDeclaration.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = Model.Declarations.FirstOrDefault(item => item.ParentScope == Model.TargetDeclaration.ParentScope &&
                                              item.IdentifierName == Model.TargetDeclaration.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertySet);

                if (setter != null)
                {
                    AdjustSignatures(setter, rewriteSession);
                    AdjustReferences(setter.References, rewriteSession);
                }

                var letter = Model.Declarations.FirstOrDefault(item => item.ParentScope == Model.TargetDeclaration.ParentScope &&
                              item.IdentifierName == Model.TargetDeclaration.IdentifierName &&
                              item.DeclarationType == DeclarationType.PropertyLet);

                if (letter != null)
                {
                    AdjustSignatures(letter, rewriteSession);
                    AdjustReferences(letter.References, rewriteSession);
                }
            }

            RewriteSignature(Model.TargetDeclaration, paramList, rewriteSession);

            foreach (var withEvents in Model.Declarations.Where(item => item.IsWithEvents && item.AsTypeName == Model.TargetDeclaration.ComponentName))
            {
                foreach (var reference in Model.Declarations.FindEventProcedures(withEvents))
                {
                    AdjustReferences(reference.References, rewriteSession);
                    AdjustSignatures(reference, rewriteSession);
                }
            }

            if (!(Model.TargetDeclaration is ModuleBodyElementDeclaration member) 
                || !(member.IsInterfaceImplementation || member.IsInterfaceMember))
            {
                return;
            }

            var implementations =
                Model.State.DeclarationFinder.FindInterfaceImplementationMembers(member.IsInterfaceMember
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
            for (var i = 0; i < Model.Parameters.Count; i++)
            {
                var param = paramList.arg()[i];
                rewriter.Replace(param, parameters.SingleOrDefault(s => s.Index == Model.Parameters[i].Index)?.Text);
            }
        }
    }
}
