using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime.Misc;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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

                // ReSharper disable once LocalizableElement
                throw new ArgumentException("Invalid declaration type", "target");
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

            QualifiedSelection? oldSelection = null;
            var pane = _vbe.ActiveCodePane;
            var module = pane.CodeModule;
            {
                if (_vbe.ActiveCodePane != null)
                {
                    oldSelection = module.GetQualifiedSelection();
                }

                RemoveVariable(target);
                UpdateSignature(target);

                if (oldSelection.HasValue)
                {
                    pane.Selection = oldSelection.Value.Selection;
                }

                _state.OnParseRequested(this);
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

        private void UpdateSignature(Declaration targetVariable)
        {
            var functionDeclaration = _declarations.FindTarget(targetVariable.QualifiedSelection, ValidDeclarationTypes);

            var proc = (dynamic)functionDeclaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = functionDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            {
                var interfaceImplementation = GetInterfaceImplementation(functionDeclaration);

                if (functionDeclaration.DeclarationType != DeclarationType.PropertyGet &&
                    functionDeclaration.DeclarationType != DeclarationType.PropertyLet &&
                    functionDeclaration.DeclarationType != DeclarationType.PropertySet)
                {
                    AddParameter(functionDeclaration, targetVariable, paramList, module);

                    if (interfaceImplementation == null) { return; }
                }

                if (functionDeclaration.DeclarationType == DeclarationType.PropertyGet ||
                    functionDeclaration.DeclarationType == DeclarationType.PropertyLet ||
                    functionDeclaration.DeclarationType == DeclarationType.PropertySet)
                {
                    UpdateProperties(functionDeclaration, targetVariable);
                }

                if (interfaceImplementation == null) { return; }

                UpdateSignature(interfaceImplementation, targetVariable);

                var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.ProjectId == interfaceImplementation.ProjectId
                                                               && item.IdentifierName == interfaceImplementation.ComponentName + "_" + interfaceImplementation.IdentifierName
                                                               && !item.Equals(functionDeclaration));

                foreach (var implementation in interfaceImplementations)
                {
                    UpdateSignature(implementation, targetVariable);
                }
            }
        }

        private void UpdateSignature(Declaration targetMethod, Declaration targetVariable)
        {
            var proc = (dynamic)targetMethod.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = targetMethod.QualifiedName.QualifiedModuleName.Component.CodeModule;
            {
                AddParameter(targetMethod, targetVariable, paramList, module);
            }
        }

        private void AddParameter(Declaration targetMethod, Declaration targetVariable, VBAParser.ArgListContext paramList, ICodeModule module)
        {
            var argList = paramList.arg();
            var lastParam = argList.LastOrDefault();

            var newContent = GetOldSignature(targetMethod);

            if (lastParam == null)
            {
                // Increase index by one because VBA is dumb enough to use 1-based indexing
                newContent = newContent.Insert(newContent.IndexOf('(') + 1, GetParameterDefinition(targetVariable));
            }
            else if (targetMethod.DeclarationType != DeclarationType.PropertyLet &&
                     targetMethod.DeclarationType != DeclarationType.PropertySet)
            {
                newContent = newContent.Replace(argList.Last().GetText(),
                    argList.Last().GetText() + ", " + GetParameterDefinition(targetVariable));
            }
            else
            {
                newContent = newContent.Replace(argList.Last().GetText(),
                    GetParameterDefinition(targetVariable) + ", " + argList.Last().GetText());
            }

            module.ReplaceLine(paramList.Start.Line, newContent);
            module.DeleteLines(paramList.Start.Line + 1, paramList.GetSelection().LineCount - 1);
        }

        private void UpdateProperties(Declaration knownProperty, Declaration targetVariable)
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
                UpdateSignature(property, targetVariable);
            }
        }

        private void RemoveVariable(Declaration target)
        {
            var module = target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            module.Remove(target);
        }

        private string GetOldSignature(Declaration target)
        {
            var rewriter = _state.GetRewriter(target.QualifiedName.QualifiedModuleName.Component);

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

            return rewriter.GetText(new Interval(firstTokenIndex, lastTokenIndex));
        }

        private Declaration GetInterfaceImplementation(Declaration target)
        {
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(target));

            if (interfaceImplementation == null) { return null; }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            return interfaceMember;
        }

        private string GetParameterDefinition(Declaration target)
        {
            return "ByVal " + target.IdentifierName + " As " + target.AsTypeName;
        }
    }
}
