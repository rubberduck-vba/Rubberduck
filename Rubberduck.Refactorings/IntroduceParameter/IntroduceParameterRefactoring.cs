using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.IntroduceParameter
{
    public class IntroduceParameterRefactoring : RefactoringBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IList<Declaration> _declarations;
        private readonly IMessageBox _messageBox;

        private static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        public IntroduceParameterRefactoring(IDeclarationFinderProvider declarationFinderProvider, IMessageBox messageBox, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService)
        {
            _declarationFinderProvider = declarationFinderProvider;
            //TODO: Refactor to use the declaration finder exclusively.
            _declarations = declarationFinderProvider.DeclarationFinder.AllUserDeclarations.ToList();
            _messageBox = messageBox;
        }

        public override void Refactor()
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            Refactor(activeSelection.Value);
        }

        public override void Refactor(QualifiedSelection selection)
        {
            var target = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable).FindVariable(selection);

            if (target == null)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            PromoteVariable(target);
        }

        public override void Refactor(Declaration target)
        {
            if (target == null || target.DeclarationType != DeclarationType.Variable)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            PromoteVariable(target);
        }

        private void PromoteVariable(Declaration target)
        {
            if (!PromptIfMethodImplementsInterface(target))
            {
                return;
            }

            if (new[] { DeclarationType.ClassModule, DeclarationType.ProceduralModule }.Contains(target.ParentDeclaration.DeclarationType))
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            UpdateSignature(target, rewriteSession);
            rewriter.Remove(target);

            rewriteSession.TryRewrite();
        }

        private bool PromptIfMethodImplementsInterface(Declaration targetVariable)
        {
            var functionDeclaration = (ModuleBodyElementDeclaration)_declarations.FindTarget(targetVariable.QualifiedSelection, ValidDeclarationTypes);

            if (functionDeclaration == null || !functionDeclaration.IsInterfaceImplementation)
            {
                return true;
            }

            var interfaceImplementation = functionDeclaration.InterfaceMemberImplemented;

            if (interfaceImplementation == null)
            {
                return true;
            }

            var message = string.Format(RubberduckUI.IntroduceParameter_PromptIfTargetIsInterface,
                functionDeclaration.IdentifierName, interfaceImplementation.ComponentName,
                interfaceImplementation.IdentifierName);

            return _messageBox.Question(message, RubberduckUI.IntroduceParameter_Caption);
        }

        private void UpdateSignature(Declaration targetVariable, IRewriteSession rewriteSession)
        {
            var functionDeclaration = (ModuleBodyElementDeclaration)_declarations.FindTarget(targetVariable.QualifiedSelection, ValidDeclarationTypes);

            var proc = (dynamic) functionDeclaration.Context;
            var paramList = (VBAParser.ArgListContext) proc.argList();

            if (functionDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                UpdateProperties(functionDeclaration, targetVariable, rewriteSession);               
            }
            else
            {
                AddParameter(functionDeclaration, targetVariable, paramList, rewriteSession);
            }

            var interfaceImplementation = functionDeclaration.InterfaceMemberImplemented;

            if (interfaceImplementation == null)
            {
                return;
            }

            UpdateSignature(interfaceImplementation, targetVariable, rewriteSession);

            var interfaceImplementations = _declarationFinderProvider.DeclarationFinder.FindInterfaceImplementationMembers(functionDeclaration.InterfaceMemberImplemented)
                .Where(member => !ReferenceEquals(member, functionDeclaration));

            foreach (var implementation in interfaceImplementations)
            {
                UpdateSignature(implementation, targetVariable, rewriteSession);
            }
        }

        private void UpdateSignature(Declaration targetMethod, Declaration targetVariable, IRewriteSession rewriteSession)
        {
            var proc = (dynamic) targetMethod.Context;
            var paramList = (VBAParser.ArgListContext) proc.argList();
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
