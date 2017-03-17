using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.IntroduceParameter
{
    public class IntroduceParameterRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
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

        public IntroduceParameterRefactoring(IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
        {
            _vbe = vbe;
            _state = state;
            _declarations = state.AllDeclarations.ToList();
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var pane = _vbe.ActiveCodePane;
            var module = pane.CodeModule;
            {
                var selection = module.GetQualifiedSelection();
                if (!selection.HasValue)
                {
                    _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                Refactor(selection.Value);
            }
        }

        public void Refactor(QualifiedSelection selection)
        {
            var target = _declarations.FindVariable(selection);

            if (target == null)
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            PromoteVariable(target);
        }

        public void Refactor(Declaration target)
        {
            if (target == null || target.DeclarationType != DeclarationType.Variable)
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            var rewriter = _state.GetRewriter(target);

            QualifiedSelection? oldSelection = null;
            var pane = _vbe.ActiveCodePane;
            var module = pane.CodeModule;
            if (_vbe.ActiveCodePane != null)
            {
                oldSelection = module.GetQualifiedSelection();
            }

            UpdateSignature(rewriter, target);
            rewriter.Remove(target);

            if (oldSelection.HasValue)
            {
                pane.Selection = oldSelection.Value.Selection;
            }
        }

        private bool PromptIfMethodImplementsInterface(Declaration targetVariable)
        {
            var functionDeclaration = _declarations.FindTarget(targetVariable.QualifiedSelection, ValidDeclarationTypes);
            var interfaceImplementation = GetInterfaceImplementation(functionDeclaration);

            if (interfaceImplementation == null)
            {
                return true;
            }

            var message = string.Format(RubberduckUI.IntroduceParameter_PromptIfTargetIsInterface,
                functionDeclaration.IdentifierName, interfaceImplementation.ComponentName,
                interfaceImplementation.IdentifierName);
            var introduceParamToInterface = _messageBox.Show(message, RubberduckUI.IntroduceParameter_Caption,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            return introduceParamToInterface != DialogResult.No;
        }

        private void UpdateSignature(IModuleRewriter rewriter, Declaration targetVariable)
        {
            var functionDeclaration = _declarations.FindTarget(targetVariable.QualifiedSelection, ValidDeclarationTypes);

            var proc = (dynamic) functionDeclaration.Context;
            var paramList = (VBAParser.ArgListContext) proc.argList();
            var interfaceImplementation = GetInterfaceImplementation(functionDeclaration);

            if (functionDeclaration.DeclarationType != DeclarationType.PropertyGet &&
                functionDeclaration.DeclarationType != DeclarationType.PropertyLet &&
                functionDeclaration.DeclarationType != DeclarationType.PropertySet)
            {
                AddParameter(functionDeclaration, targetVariable, paramList);

                if (interfaceImplementation == null)
                {
                    return;
                }
            }

            if (functionDeclaration.DeclarationType == DeclarationType.PropertyGet ||
                functionDeclaration.DeclarationType == DeclarationType.PropertyLet ||
                functionDeclaration.DeclarationType == DeclarationType.PropertySet)
            {
                UpdateProperties(rewriter, functionDeclaration, targetVariable);
            }

            if (interfaceImplementation == null)
            {
                return;
            }

            UpdateSignature(rewriter, interfaceImplementation, targetVariable);

            var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                .Where(item => item.ProjectId == interfaceImplementation.ProjectId
                               &&
                               item.IdentifierName ==
                               interfaceImplementation.ComponentName + "_" + interfaceImplementation.IdentifierName
                               && !item.Equals(functionDeclaration));

            foreach (var implementation in interfaceImplementations)
            {
                UpdateSignature(rewriter, implementation, targetVariable);
            }
        }

        private void UpdateSignature(IModuleRewriter rewriter, Declaration targetMethod, Declaration targetVariable)
        {
            var proc = (dynamic) targetMethod.Context;
            var paramList = (VBAParser.ArgListContext) proc.argList();
            AddParameter(targetMethod, targetVariable, paramList);
        }

        private void AddParameter(Declaration targetMethod, Declaration targetVariable, VBAParser.ArgListContext paramList)
        {
            var rewriter = _state.GetRewriter(targetMethod);

            var argList = paramList.arg();
            var newParameter = Tokens.ByVal + " " + targetVariable.IdentifierName + " "+ Tokens.As + " " + targetVariable.AsTypeName;

            if (!argList.Any())
            {
                rewriter.InsertAtIndex(newParameter, paramList.RPAREN().Symbol.TokenIndex);
            }
            else if (targetMethod.DeclarationType != DeclarationType.PropertyLet &&
                     targetMethod.DeclarationType != DeclarationType.PropertySet)
            {
                rewriter.InsertAtIndex($", {newParameter}", paramList.RPAREN().Symbol.TokenIndex);
            }
            else
            {
                var lastParam = argList.Last();
                rewriter.InsertAtIndex($"{newParameter}, ", lastParam.Start.TokenIndex);
            }
        }

        private void UpdateProperties(IModuleRewriter rewriter, Declaration knownProperty, Declaration targetVariable)
        {
            var propertyGet = _declarations.FirstOrDefault(d =>
                    d.DeclarationType == DeclarationType.PropertyGet &&
                    d.Scope == knownProperty.Scope &&
                    d.IdentifierName == knownProperty.IdentifierName);

            var propertyLet = _declarations.FirstOrDefault(d =>
                    d.DeclarationType == DeclarationType.PropertyLet &&
                    d.Scope == knownProperty.Scope &&
                    d.IdentifierName == knownProperty.IdentifierName);

            var propertySet = _declarations.FirstOrDefault(d =>
                    d.DeclarationType == DeclarationType.PropertySet &&
                    d.Scope == knownProperty.Scope &&
                    d.IdentifierName == knownProperty.IdentifierName);

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
                UpdateSignature(rewriter, property, targetVariable);
            }
        }

        private string GetOldSignature(IModuleRewriter rewriter, Declaration target)
        {
            var context = target.Context;
            var firstTokenIndex = context.Start.TokenIndex;
            var lastTokenIndex = -1; // will blow up if this code runs for any context other than below

            var subStmtContext = context as VBAParser.SubStmtContext;
            if (subStmtContext != null)
            {
                lastTokenIndex = subStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            var functionStmtContext = context as VBAParser.FunctionStmtContext;
            if (functionStmtContext != null)
            {
                lastTokenIndex = functionStmtContext.asTypeClause() != null
                    ? functionStmtContext.asTypeClause().Stop.TokenIndex
                    : functionStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            var propertyGetStmtContext = context as VBAParser.PropertyGetStmtContext;
            if (propertyGetStmtContext != null)
            {
                lastTokenIndex = propertyGetStmtContext.asTypeClause() != null
                    ? propertyGetStmtContext.asTypeClause().Stop.TokenIndex
                    : propertyGetStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            var propertyLetStmtContext = context as VBAParser.PropertyLetStmtContext;
            if (propertyLetStmtContext != null)
            {
                lastTokenIndex = propertyLetStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            var propertySetStmtContext = context as VBAParser.PropertySetStmtContext;
            if (propertySetStmtContext != null)
            {
                lastTokenIndex = propertySetStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            return rewriter.GetText(firstTokenIndex, lastTokenIndex);
        }

        private Declaration GetInterfaceImplementation(Declaration target)
        {
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(target));

            if (interfaceImplementation == null) { return null; }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            return interfaceMember;
        }
    }
}
