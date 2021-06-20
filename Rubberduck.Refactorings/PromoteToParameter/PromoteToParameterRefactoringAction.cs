using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;


namespace Rubberduck.Refactorings.PromoteToParameter
{
    public class PromoteToParameterRefactoringAction : RefactoringActionBase<PromoteToParameterModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public PromoteToParameterRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        protected override void Refactor(PromoteToParameterModel model, IRewriteSession rewriteSession)
        {
            var target = model.Target;
            UpdateSignature(target, model.EnclosingMember, rewriteSession);
            RemoveTarget(target, rewriteSession);
        }

        private static void RemoveTarget(Declaration target, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
            rewriter.Remove(target);
        }

        private void UpdateSignature(Declaration targetVariable, ModuleBodyElementDeclaration enclosingMember, IRewriteSession rewriteSession)
        {
            var paramList = enclosingMember.Context.GetChild<VBAParser.ArgListContext>();

            if (enclosingMember.DeclarationType.HasFlag(DeclarationType.Property))
            {
                UpdateProperties(enclosingMember, targetVariable, rewriteSession);
            }
            else
            {
                AddParameter(enclosingMember, targetVariable, paramList, rewriteSession);
            }

            var interfaceImplementation = enclosingMember.InterfaceMemberImplemented;

            if (interfaceImplementation == null)
            {
                return;
            }

            UpdateSignature(interfaceImplementation, targetVariable, rewriteSession);

            var interfaceImplementations = _declarationFinderProvider.DeclarationFinder.FindInterfaceImplementationMembers(enclosingMember.InterfaceMemberImplemented)
                .Where(member => !ReferenceEquals(member, enclosingMember));

            foreach (var implementation in interfaceImplementations)
            {
                UpdateSignature(implementation, targetVariable, rewriteSession);
            }
        }

        private void UpdateSignature(Declaration targetMethod, Declaration targetVariable, IRewriteSession rewriteSession)
        {
            var proc = (dynamic)targetMethod.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            AddParameter(targetMethod, targetVariable, paramList, rewriteSession);
        }

        private void AddParameter(Declaration targetMethod, Declaration targetVariable, VBAParser.ArgListContext paramList, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(targetMethod.QualifiedModuleName);

            var argList = paramList.arg();
            var newParameter = $"{Tokens.ByVal} {targetVariable.IdentifierName} {Tokens.As} {targetVariable.AsTypeName}";

            if (!argList.Any())
            {
                rewriter.InsertBefore(paramList.RPAREN().Symbol.TokenIndex, newParameter);
            }
            else if (targetMethod.DeclarationType != DeclarationType.PropertyLet &&
                     targetMethod.DeclarationType != DeclarationType.PropertySet)
            {
                rewriter.InsertBefore(paramList.RPAREN().Symbol.TokenIndex, $", {newParameter}");
            }
            else
            {
                var lastParam = argList.Last();
                rewriter.InsertBefore(lastParam.Start.TokenIndex, $"{newParameter}, ");
            }
        }

        private void UpdateProperties(Declaration knownProperty, Declaration targetVariable, IRewriteSession rewriteSession)
        {
            var declarationFinder = _declarationFinderProvider.DeclarationFinder;

            var propertyGet = declarationFinder.UserDeclarations(DeclarationType.PropertyGet)
                .FirstOrDefault(d =>
                    d.QualifiedModuleName.Equals(knownProperty.QualifiedModuleName)
                    && d.IdentifierName == knownProperty.IdentifierName);

            var propertyLet = declarationFinder.UserDeclarations(DeclarationType.PropertyLet)
                .FirstOrDefault(d =>
                    d.QualifiedModuleName.Equals(knownProperty.QualifiedModuleName)
                    && d.IdentifierName == knownProperty.IdentifierName);

            var propertySet = declarationFinder.UserDeclarations(DeclarationType.PropertySet)
                .FirstOrDefault(d =>
                    d.QualifiedModuleName.Equals(knownProperty.QualifiedModuleName)
                    && d.IdentifierName == knownProperty.IdentifierName);

            var properties = new List<Declaration>();

            if (propertyGet != null)
            {
                properties.Add(propertyGet);
            }

            if (propertyLet != null)
            {
                properties.Add(propertyLet);
            }

            if (propertySet != null)
            {
                properties.Add(propertySet);
            }

            foreach (var property in
                    properties.OrderByDescending(o => o.Selection.StartLine)
                        .ThenByDescending(t => t.Selection.StartColumn))
            {
                UpdateSignature(property, targetVariable, rewriteSession);
            }
        }
    }
}