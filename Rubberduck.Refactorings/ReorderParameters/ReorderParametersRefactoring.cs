using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Rewriter;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
 using Rubberduck.Refactorings.Exceptions;
 using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersRefactoring : InteractiveRefactoringBase<IReorderParametersPresenter, ReorderParametersModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ReorderParametersRefactoring(IDeclarationFinderProvider declarationFinderProvider, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            return _declarationFinderProvider.DeclarationFinder
                .AllUserDeclarations
                .FindTarget(targetSelection, ValidDeclarationTypes);
        }

        protected override ReorderParametersModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            var model = DerivedTarget(new ReorderParametersModel(target));

            return model;
        }

        private ReorderParametersModel DerivedTarget(ReorderParametersModel model)
        {
            var preliminarymodel = ResolvedInterfaceMemberTarget(model)
                                   ?? ResolvedEventTarget(model)
                                   ?? model;
            return ResolvedGetterTarget(preliminarymodel) ?? preliminarymodel;
        }

        private static ReorderParametersModel ResolvedInterfaceMemberTarget(ReorderParametersModel model)
        {
            var declaration = model.TargetDeclaration;
            if (!(declaration is ModuleBodyElementDeclaration member) || !member.IsInterfaceImplementation)
            {
                return null;
            }

            model.IsInterfaceMemberRefactoring = true;
            model.TargetDeclaration = member.InterfaceMemberImplemented;

            return model;
        }

        private ReorderParametersModel ResolvedEventTarget(ReorderParametersModel model)
        {
            foreach (var events in _declarationFinderProvider
                .DeclarationFinder
                .UserDeclarations(DeclarationType.Event))
            {
                if (_declarationFinderProvider.DeclarationFinder
                    .AllUserDeclarations
                    .FindHandlersForEvent(events)
                    .Any(reference => Equals(reference.Item2, model.TargetDeclaration)))
                {
                    model.IsEventRefactoring = true;
                    model.TargetDeclaration = events;
                    return model;
                }
            }
            return null;
        }

        private ReorderParametersModel ResolvedGetterTarget(ReorderParametersModel model)
        {
            var target = model.TargetDeclaration;
            if (target == null || !target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                return null;
            }

            if (target.DeclarationType == DeclarationType.PropertyGet)
            {
                model.IsPropertyRefactoringWithGetter = true;
                return model;
            }


            var getter = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.PropertyGet)
                .FirstOrDefault(item => item.Scope == target.Scope
                                        && item.IdentifierName == target.IdentifierName);

            if (getter == null)
            {
                return null;
            }

            model.IsPropertyRefactoringWithGetter = true;
            model.TargetDeclaration = getter;

            return model;
        }

        protected override void RefactorImpl(ReorderParametersModel model)
        {
            if (!model.Parameters.Where((param, index) => param.Index != index).Any())
            {
                //This is not an error: the user chose to leave everything as-is.
                return;
            }

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            AdjustReferences(model, model.TargetDeclaration.References, rewriteSession);
            AdjustSignatures(model, rewriteSession);
            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private void AdjustReferences(ReorderParametersModel model, IEnumerable<IdentifierReference> references, IRewriteSession rewriteSession)
        {
            foreach (var reference in references.Where(item => item.Context != model.TargetDeclaration.Context))
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
                RewriteCall(model, argumentList, module, rewriteSession);
            }
        }

        private void RewriteCall(ReorderParametersModel model, VBAParser.ArgumentListContext argList, QualifiedModuleName module, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(module);

            var args = argList.argument().Select((s, i) => new { Index = i, Text = s.GetText() }).ToList();
            for (var i = 0; i < model.Parameters.Count; i++)
            {
                if (argList.argument().Length <= i)
                {
                    break;
                }

                var arg = argList.argument()[i];
                rewriter.Replace(arg, args.Single(s => s.Index == model.Parameters[i].Index).Text);
            }
        }

        private void AdjustSignatures(ReorderParametersModel model, IRewriteSession rewriteSession)
        {
            var proc = (dynamic)model.TargetDeclaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();

            // if we are reordering a property getter, check if we need to reorder a letter/setter too
            if (model.IsPropertyRefactoringWithGetter)
            {
                var setter = _declarationFinderProvider.DeclarationFinder
                    .UserDeclarations(DeclarationType.PropertySet)
                    .FirstOrDefault(item => item.ParentScope == model.TargetDeclaration.ParentScope
                                                && item.IdentifierName == model.TargetDeclaration.IdentifierName);

                if (setter != null)
                {
                    AdjustSignatures(model, setter, rewriteSession);
                    AdjustReferences(model, setter.References, rewriteSession);
                }

                var letter = _declarationFinderProvider.DeclarationFinder
                    .UserDeclarations(DeclarationType.PropertyLet)
                    .FirstOrDefault(item => item.ParentScope == model.TargetDeclaration.ParentScope 
                                                && item.IdentifierName == model.TargetDeclaration.IdentifierName);

                if (letter != null)
                {
                    AdjustSignatures(model, letter, rewriteSession);
                    AdjustReferences(model, letter.References, rewriteSession);
                }
            }

            RewriteSignature(model, model.TargetDeclaration, paramList, rewriteSession);

            foreach (var withEvents in _declarationFinderProvider.DeclarationFinder
                .AllUserDeclarations
                .Where(item => item.IsWithEvents && item.AsTypeName == model.TargetDeclaration.ComponentName))
            {
                foreach (var reference in _declarationFinderProvider.DeclarationFinder
                    .AllUserDeclarations
                    .FindEventProcedures(withEvents))
                {
                    AdjustReferences(model, reference.References, rewriteSession);
                    AdjustSignatures(model, reference, rewriteSession);
                }
            }

            if (!(model.TargetDeclaration is ModuleBodyElementDeclaration member) 
                || !(member.IsInterfaceImplementation || member.IsInterfaceMember))
            {
                return;
            }

            var implementations =
                _declarationFinderProvider.DeclarationFinder.FindInterfaceImplementationMembers(member.IsInterfaceMember
                    ? member
                    : member.InterfaceMemberImplemented);

            foreach (var interfaceImplentation in implementations)
            {
                AdjustReferences(model, interfaceImplentation.References, rewriteSession);
                AdjustSignatures(model, interfaceImplentation, rewriteSession);
            }
        }

        private static void AdjustSignatures(ReorderParametersModel model, Declaration declaration, IRewriteSession rewriteSession)
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

            RewriteSignature(model, declaration, paramList, rewriteSession);
        }

        private static void RewriteSignature(ReorderParametersModel model, Declaration target, VBAParser.ArgListContext paramList, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            var parameters = paramList.arg().Select((s, i) => new { Index = i, Text = s.GetText() }).ToList();
            for (var i = 0; i < model.Parameters.Count; i++)
            {
                var param = paramList.arg()[i];
                rewriter.Replace(param, parameters.SingleOrDefault(s => s.Index == model.Parameters[i].Index)?.Text);
            }
        }

        public static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };
    }
}
