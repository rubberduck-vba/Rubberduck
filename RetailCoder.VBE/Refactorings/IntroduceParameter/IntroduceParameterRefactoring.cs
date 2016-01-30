using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime.Misc;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.IntroduceParameter
{
    public class IntroduceParameter : IRefactoring
    {
        private readonly RubberduckParserState _parseResult;
        private readonly IList<Declaration> _declarations;
        private readonly IActiveCodePaneEditor _editor;
        private readonly IMessageBox _messageBox;

        private static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        public IntroduceParameter(RubberduckParserState parseResult, IActiveCodePaneEditor editor, IMessageBox messageBox)
        {
            _parseResult = parseResult;
            _declarations = parseResult.AllDeclarations.ToList();
            _editor = editor;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var selection = _editor.GetSelection();

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

            PromoteVariable(target);
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Invalid declaration type");
            }

            PromoteVariable(target);
        }

        private void PromoteVariable(Declaration target)
        {
            if (!PromptIfMethodImplementsInterface(target))
            {
                return;
            }

            if (IsContainedInClassOrModule(target))
            {
                return;
            }

            RemoveVariable(target);
            UpdateSignature(target);
        }

        private bool IsContainedInClassOrModule(Declaration target)
        {
            return target.ParentDeclaration != null && IsContainedInClassOrModule(target.ParentDeclaration);
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

            AddParameter(functionDeclaration, targetVariable, paramList, module);

            if (functionDeclaration.DeclarationType == DeclarationType.PropertyGet ||
                functionDeclaration.DeclarationType == DeclarationType.PropertyLet ||
                functionDeclaration.DeclarationType == DeclarationType.PropertySet)
            {
                UpdateProperties(functionDeclaration);
            }

            var interfaceImplementation = GetInterfaceImplementation(functionDeclaration);
            if (interfaceImplementation == null)
            {
                return;
            }
            UpdateSignature(interfaceImplementation, targetVariable);

            var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                                                    .Where(item => item.Project.Equals(interfaceImplementation.Project)
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
        }

        private void UpdateProperties(Declaration target)
        {
            var propertyGet = _declarations.FirstOrDefault(d =>
                    d.DeclarationType == DeclarationType.PropertyGet &&
                    d.Scope == target.Scope &&
                    d.IdentifierName == target.IdentifierName);

            var propertyLet = _declarations.FirstOrDefault(d =>
                    d.DeclarationType == DeclarationType.PropertyLet &&
                    d.Scope == target.Scope &&
                    d.IdentifierName == target.IdentifierName);

            var propertySet = _declarations.FirstOrDefault(d =>
                    d.DeclarationType == DeclarationType.PropertySet &&
                    d.Scope == target.Scope &&
                    d.IdentifierName == target.IdentifierName);

            if (target.DeclarationType != DeclarationType.PropertyGet && propertyGet != null)
            {
                UpdateSignature(propertyGet);
            }

            if (target.DeclarationType != DeclarationType.PropertyLet && propertyLet != null)
            {
                UpdateSignature(propertyLet);
            }

            if (target.DeclarationType != DeclarationType.PropertySet && propertySet != null)
            {
                UpdateSignature(propertySet);
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

            var oldLines = _editor.GetLines(selection);

            var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                .Remove(selection.StartColumn, declarationText.Length);

            if (multipleDeclarations)
            {
                selection = target.GetVariableStmtContextSelection();
                newLines = RemoveExtraComma(_editor.GetLines(selection).Replace(oldLines, newLines),
                    target.CountOfDeclarationsInStatement(), target.IndexOfVariableDeclarationInStatement());
            }

            _editor.DeleteLines(selection);
            var newLinesWithoutEmptyLines = newLines
                    .Split(new[] {" _" + Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                    .Where(l => l.Trim() != string.Empty).ToList();

            if (newLinesWithoutEmptyLines.Any())
            {
                _editor.InsertLines(selection.StartLine, string.Join(" _" + Environment.NewLine, newLinesWithoutEmptyLines));
            }
        }

        private string GetOldSignature(Declaration target)
        {
            var rewriter = _parseResult.GetRewriter(target.QualifiedName.QualifiedModuleName.Component);

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

            var declareStmtContext = context as VBAParser.DeclareStmtContext;
            if (declareStmtContext != null)
            {
                lastTokenIndex = declareStmtContext.STRINGLITERAL().Last().Symbol.TokenIndex;
                if (declareStmtContext.argList() != null)
                {
                    lastTokenIndex = declareStmtContext.argList().RPAREN().Symbol.TokenIndex;
                }
                if (declareStmtContext.asTypeClause() != null)
                {
                    lastTokenIndex = declareStmtContext.asTypeClause().Stop.TokenIndex;
                }
            }

            var eventStmtContext = context as VBAParser.EventStmtContext;
            if (eventStmtContext != null)
            {
                lastTokenIndex = eventStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            return rewriter.GetText(new Interval(firstTokenIndex, lastTokenIndex));
        }

        private Declaration GetInterfaceImplementation(Declaration target)
        {
            var declaration = target;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));

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
            if (target == null) { return null; }

            return "ByVal " + target.IdentifierName + " As " + target.AsTypeName;
        }
    }
}
