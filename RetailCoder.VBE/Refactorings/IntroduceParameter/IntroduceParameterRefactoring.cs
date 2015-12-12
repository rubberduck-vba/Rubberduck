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
        private Declaration _targetDeclaration;
        private readonly IMessageBox _messageBox;

        private static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        public IntroduceParameter (RubberduckParserState parseResult, IActiveCodePaneEditor editor, IMessageBox messageBox)
        {
            _parseResult = parseResult;
            _declarations = parseResult.AllDeclarations.ToList();
            _editor = editor;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            if (_targetDeclaration == null)
            {
                _messageBox.Show("Invalid selection...");   // todo: write a better message and localize it
                return;
            }

            UpdateSignature();
            RemoveVariable();
        }

        public void Refactor (QualifiedSelection target)
        {
            _targetDeclaration = _declarations.FindSelection(target, new [] {DeclarationType.Variable});

            _editor.SetSelection(target);
            Refactor();
        }

        public void Refactor (Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Invalid declaration type");
            }

            _targetDeclaration = target;
            _editor.SetSelection(target.QualifiedSelection);
            Refactor();
        }

        private void UpdateSignature()
        {
            var functionDeclaration = _declarations.FindSelection(_targetDeclaration.QualifiedSelection, ValidDeclarationTypes);

            var proc = (dynamic)functionDeclaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = functionDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;

            AddParameter(functionDeclaration, paramList, module);
        }

        private void AddParameter(Declaration target, VBAParser.ArgListContext paramList, CodeModule module)
        {
            var argList = paramList.arg();

            var newContent = GetOldSignature(target);

            var lastParam = argList.LastOrDefault();

            if (lastParam == null)
            {
                // Increase index by one because VBA is dumb enough to use 1-based indexing
                newContent = newContent.Insert(newContent.IndexOf('(') + 1, GetParameterDefinition());
            }
            else
            {
                newContent = newContent.Replace(argList.Last().GetText(),
                    argList.Last().GetText() + ", " + GetParameterDefinition());
            }

            module.ReplaceLine(paramList.Start.Line, newContent);
        }

        private void RemoveVariable()
        {
            Selection selection;
            var declarationText = _targetDeclaration.Context.GetText();
            var multipleDeclarations = HasMultipleDeclarationsInStatement();

            var variableStmtContext = GetVariableStmtContext();

            if (!multipleDeclarations)
            {
                declarationText = variableStmtContext.GetText();
                selection = GetVariableStmtContextSelection();
            }
            else
            {
                selection = new Selection(_targetDeclaration.Context.Start.Line, _targetDeclaration.Context.Start.Column,
                    _targetDeclaration.Context.Stop.Line, _targetDeclaration.Context.Stop.Column);
            }

            var oldLines = _editor.GetLines(selection);

            var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                .Remove(selection.StartColumn, declarationText.Length);

            if (multipleDeclarations)
            {
                selection = GetVariableStmtContextSelection();
                newLines = RemoveExtraComma(_editor.GetLines(selection).Replace(oldLines, newLines));
            }

            _editor.DeleteLines(selection);
            _editor.InsertLines(selection.StartLine, newLines);
        }

        private string GetOldSignature (Declaration target)
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

        private Selection GetVariableStmtContextSelection()
        {
            var statement = GetVariableStmtContext();

            return new Selection(statement.Start.Line, statement.Start.Column,
                    statement.Stop.Line, statement.Stop.Column);
        }

        private VBAParser.VariableStmtContext GetVariableStmtContext()
        {
            var statement = _targetDeclaration.Context.Parent.Parent as VBAParser.VariableStmtContext;
            if (statement == null)
            {
                throw new MissingMemberException("Statement not found");
            }

            return statement;
        }

        private string RemoveExtraComma(string str)
        {
            if (str.Count(c => c == ',') == 1)
            {
                return str.Remove(str.IndexOf(','), 1);
            }

            var significantCharacterAfterComma = false;

            for (var index = 0; index < str.Length; index++)
            {
                if (!char.IsWhiteSpace(str[index]) && str[index] != '_' && str[index] != ',')
                {
                    significantCharacterAfterComma = true;
                }
                if (str[index] == ',')
                {
                    significantCharacterAfterComma = false;
                }

                if (!significantCharacterAfterComma && str[index] == ',')
                {
                    return str.Remove(index, 1);
                }
            }

            return str;
        }

        private bool HasMultipleDeclarationsInStatement()
        {
            var statement = _targetDeclaration.Context.Parent as VBAParser.VariableListStmtContext;

            if (statement == null) { return false; }

            return statement.children.Count(i => i is VBAParser.VariableSubStmtContext) > 1;
        }

        private string GetParameterDefinition()
        {
            if (_targetDeclaration == null) { return null; }

            return "ByVal " + _targetDeclaration.IdentifierName + " As " + _targetDeclaration.AsTypeName;
        }
    }
}
