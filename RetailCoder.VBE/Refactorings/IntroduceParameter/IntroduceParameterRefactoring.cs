using System;
using System.Collections.Generic;
using System.Linq;
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
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection);
                return;
            }

            Refactor(selection.Value);
        }

        public void Refactor(QualifiedSelection selection)
        {
            var target = FindSelection(selection);

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
            RemoveVariable(target);
            UpdateSignature(target);
        }

        private void UpdateSignature(Declaration targetVariable)
        {
            var functionDeclaration = _declarations.FindSelection(targetVariable.QualifiedSelection, ValidDeclarationTypes);

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

            var interfaceDeclaration = GetInterfaceImplementation(functionDeclaration);
            if (interfaceDeclaration != null)
            {
                UpdateSignature(interfaceDeclaration, targetVariable);

                var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                                            .Where(item => item.Project.Equals(interfaceDeclaration.Project) &&
                                                   item.IdentifierName == interfaceDeclaration.ComponentName + "_" + interfaceDeclaration.IdentifierName);

                foreach (var interfaceImplementation in interfaceImplementations)
                {
                    UpdateSignature(interfaceImplementation, targetVariable);
                }
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
            var multipleDeclarations = HasMultipleDeclarationsInStatement(target);

            var variableStmtContext = GetVariableStmtContext(target);

            if (!multipleDeclarations)
            {
                declarationText = variableStmtContext.GetText();
                selection = GetVariableStmtContextSelection(target);
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
                selection = GetVariableStmtContextSelection(target);
                newLines = RemoveExtraComma(_editor.GetLines(selection).Replace(oldLines, newLines));
            }

            _editor.DeleteLines(selection);
            _editor.InsertLines(selection.StartLine, newLines);
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

        private Selection GetVariableStmtContextSelection(Declaration target)
        {
            var statement = GetVariableStmtContext(target);

            return new Selection(statement.Start.Line, statement.Start.Column,
                    statement.Stop.Line, statement.Stop.Column);
        }

        private VBAParser.VariableStmtContext GetVariableStmtContext(Declaration target)
        {
            var statement = target.Context.Parent.Parent as VBAParser.VariableStmtContext;
            if (statement == null)
            {
                throw new MissingMemberException("Statement not found");
            }

            return statement;
        }

        private Declaration GetInterfaceImplementation(Declaration target)
        {
            var declaration = target;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));

            if (interfaceImplementation == null) { return null; }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            return interfaceMember;
        }

        private string RemoveExtraComma(string str)
        {
            if (str.Count(c => c == ',') == 1)
            {
                return str.Remove(str.IndexOf(','), 1);
            }

            var significantCharacterAfterComma = false;

            for (var index = str.IndexOf("Dim", StringComparison.Ordinal) + 3; index < str.Length; index++)
            {
                if (!significantCharacterAfterComma && str[index] == ',')
                {
                    return str.Remove(index, 1);
                }

                if (!char.IsWhiteSpace(str[index]) && str[index] != '_' && str[index] != ',')
                {
                    significantCharacterAfterComma = true;
                }

                if (str[index] == ',')
                {
                    significantCharacterAfterComma = false;
                }
            }

            return str.Remove(str.LastIndexOf(','), 1);
        }

        private bool HasMultipleDeclarationsInStatement(Declaration target)
        {
            var statement = target.Context.Parent as VBAParser.VariableListStmtContext;

            if (statement == null) { return false; }

            return statement.children.Count(i => i is VBAParser.VariableSubStmtContext) > 1;
        }

        private string GetParameterDefinition(Declaration target)
        {
            if (target == null) { return null; }

            return "ByVal " + target.IdentifierName + " As " + target.AsTypeName;
        }

        private Declaration FindSelection(QualifiedSelection selection)
        {
            var target = _declarations
                .Where(item => !item.IsBuiltIn)
                .FirstOrDefault(item => item.IsSelected(selection) && item.DeclarationType == DeclarationType.Variable
                                     || item.References.Any(r => r.IsSelected(selection) &&
                                        r.Declaration.DeclarationType == DeclarationType.Variable));

            if (target != null) { return target; }

            var targets = _declarations
                .Where(item => !item.IsBuiltIn
                               && item.ComponentName == selection.QualifiedName.ComponentName
                               && item.DeclarationType == DeclarationType.Variable);

            foreach (var declaration in targets)
            {
                var declarationSelection = new Selection(declaration.Context.Start.Line,
                                                    declaration.Context.Start.Column,
                                                    declaration.Context.Stop.Line,
                                                    declaration.Context.Stop.Column + declaration.Context.Stop.Text.Length);

                if (declarationSelection.Contains(selection.Selection) ||
                    !HasMultipleDeclarationsInStatement(declaration) && GetVariableStmtContextSelection(declaration).Contains(selection.Selection))
                {
                    return declaration;
                }

                var reference =
                    declaration.References.FirstOrDefault(r => r.Selection.Contains(selection.Selection));

                if (reference != null)
                {
                    return reference.Declaration;
                }
            }
            return null;
        }
    }
}
