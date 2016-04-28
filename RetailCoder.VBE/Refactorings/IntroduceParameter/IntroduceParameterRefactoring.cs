using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime.Misc;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.IntroduceParameter
{
    public class IntroduceParameterRefactoring : IRefactoring
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _parserState;
        private readonly IMessageBox _messageBox;

        private IList<Declaration> _declarations;

        private static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        public IntroduceParameterRefactoring(VBE vbe, RubberduckParserState parserState, IMessageBox messageBox)
        {
            _vbe = vbe;
            _parserState = parserState;
            _messageBox = messageBox;
        }

        public bool CanExecute(QualifiedSelection selection)
        {
            _declarations = _parserState.AllUserDeclarations.ToList();

            var target = _declarations.FindVariable(selection);
            return target != null;
        }

        public void Refactor()
        {
            var selection = _vbe.ActiveCodePane.CodeModule.GetSelection();

            if (!selection.HasValue)
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            Refactor(selection.Value);
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

            RemoveVariable(target);
            UpdateSignature(target);
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

        private void UpdateSignature(Declaration targetMethod, Declaration targetVariable)
        {
            var proc = (dynamic)targetMethod.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = targetMethod.QualifiedName.QualifiedModuleName.Component.CodeModule;

            AddParameter(targetMethod, targetVariable, paramList, module);
        }

        private void AddParameter(Declaration targetMethod, Declaration targetVariable, VBAParser.ArgListContext paramList, CodeModule module)
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
            Selection selection;
            var declarationText = target.Context.GetText();
            var multipleDeclarations = target.HasMultipleDeclarationsInStatement();

            var variableStmtContext = target.GetVariableStmtContext();

            if (!multipleDeclarations)
            {
                declarationText = variableStmtContext.GetText();
                selection = target.GetVariableStmtContextSelection();
            }
            else
            {
                selection = new Selection(target.Context.Start.Line, target.Context.Start.Column,
                    target.Context.Stop.Line, target.Context.Stop.Column);
            }

            var oldLines = _vbe.ActiveCodePane.CodeModule.GetLines(selection);

            var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                .Remove(selection.StartColumn, declarationText.Length);

            if (multipleDeclarations)
            {
                selection = target.GetVariableStmtContextSelection();
                newLines = RemoveExtraComma(_vbe.ActiveCodePane.CodeModule.GetLines(selection).Replace(oldLines, newLines),
                    target.CountOfDeclarationsInStatement(), target.IndexOfVariableDeclarationInStatement());
            }

            var newLinesWithoutExcessSpaces = newLines.Split(new[] {Environment.NewLine}, StringSplitOptions.None);
            for (var i = 0; i < newLinesWithoutExcessSpaces.Length; i++)
            {
                newLinesWithoutExcessSpaces[i] = newLinesWithoutExcessSpaces[i].RemoveExtraSpacesLeavingIndentation();
            }

            for (var i = newLinesWithoutExcessSpaces.Length - 1; i >= 0; i--)
            {
                if (newLinesWithoutExcessSpaces[i].Trim() == string.Empty)
                {
                    continue;
                }

                if (newLinesWithoutExcessSpaces[i].EndsWith(" _"))
                {
                    newLinesWithoutExcessSpaces[i] =
                        newLinesWithoutExcessSpaces[i].Remove(newLinesWithoutExcessSpaces[i].Length - 2);
                }
                break;
            }

            _vbe.ActiveCodePane.CodeModule.DeleteLines(selection);
            _vbe.ActiveCodePane.CodeModule.InsertLines(selection.StartLine, string.Join(Environment.NewLine, newLinesWithoutExcessSpaces));
        }

        private string GetOldSignature(Declaration target)
        {
            var rewriter = _parserState.GetRewriter(target.QualifiedName.QualifiedModuleName.Component);

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

        private string RemoveExtraComma(string str, int numParams, int indexRemoved)
        {
            // Example use cases for this method (fields and variables):
            // Dim fizz as Boolean, dizz as Double
            // Private fizz as Boolean, dizz as Double
            // Public fizz as Boolean, _
            //        dizz as Double
            // Private fizz as Boolean _
            //         , dizz as Double _
            //         , iizz as Integer

            // Before this method is called, the parameter to be removed has 
            // already been removed.  This means 'str' will look like:
            // Dim fizz as Boolean, 
            // Private , dizz as Double
            // Public fizz as Boolean, _
            //        
            // Private  _
            //         , dizz as Double _
            //         , iizz as Integer

            // This method is responsible for removing the redundant comma
            // and returning a string similar to:
            // Dim fizz as Boolean
            // Private dizz as Double
            // Public fizz as Boolean _
            //        
            // Private  _
            //          dizz as Double _
            //         , iizz as Integer

            var commaToRemove = numParams == indexRemoved ? indexRemoved - 1 : indexRemoved;

            return str.Remove(str.NthIndexOf(',', commaToRemove), 1);
        }

        private string GetParameterDefinition(Declaration target)
        {
            return "ByVal " + target.IdentifierName + " As " + target.AsTypeName;
        }
    }
}
