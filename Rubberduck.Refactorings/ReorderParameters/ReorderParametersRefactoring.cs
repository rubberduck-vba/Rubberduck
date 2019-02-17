﻿using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Rewriter;
using System.Collections.Generic;
using System.Linq;
 using Rubberduck.Parsing.VBA;
 using Rubberduck.Refactorings.Exceptions;
 using Rubberduck.Refactorings.Exceptions.ReorderParameters;
 using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersRefactoring : InteractiveRefactoringBase<IReorderParametersPresenter, ReorderParametersModel>
    {
        private readonly RubberduckParserState _state;

        public ReorderParametersRefactoring(RubberduckParserState state, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _state = state;
        }

        public override void Refactor(QualifiedSelection target)
        {
            CheckAndRefactor(InitializeModel(target));
        }

        private ReorderParametersModel InitializeModel(QualifiedSelection targetSelection)
        {
            return new ReorderParametersModel(_state, targetSelection);
        }

        private void CheckAndRefactor(ReorderParametersModel model)
        {
            if (model == null)
            {
                throw new InvalidRefactoringModelException();
            }

            if (model.TargetDeclaration == null)
            {
                throw new TargetDeclarationIsNullException(null);
            }

            Refactor(model);
        }

        public override void Refactor(Declaration target)
        {
            CheckAndRefactor(InitializeModel(target));
        }

        protected ReorderParametersModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException(target);
            }

            if (!ReorderParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            return InitializeModel(target.QualifiedSelection);
        }

        protected override void RefactorImpl(IReorderParametersPresenter presenter)
        {
            if (!Model.Parameters.Where((param, index) => param.Index != index).Any())
            {
                //This is not an error: the user chose to leave everything as-is.
                return;
            }

            if (!IsValidParamOrder())
            {
                throw new InvalidParameterOrderException();
            }

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            AdjustReferences(Model.TargetDeclaration.References, rewriteSession);
            AdjustSignatures(rewriteSession);
            rewriteSession.TryRewrite();
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
                        throw new OptionalParameterNotAtTheEndException();
                    }
                }
            }

            var indexOfParamArray = Model.Parameters.FindIndex(param => param.IsParamArray);
            if (indexOfParamArray >= 0 && indexOfParamArray != Model.Parameters.Count - 1)
            {
                throw new ParamArrayIsNotLastParameterException();
            }
            
            return true;
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
