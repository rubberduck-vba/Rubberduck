using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.IntroduceParameter
{
    public class PromoteLocalToParameterRefactoring : IRefactoring
    {
        private readonly IList<Declaration> _declarations;
        private readonly IActiveCodePaneEditor _editor;
        private Declaration _targetDeclaration;
        private readonly IMessageBox _messageBox;

        public PromoteLocalToParameterRefactoring (RubberduckParserState parseResult, IActiveCodePaneEditor editor, IMessageBox messageBox)
        {
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

        private void AddParameter()
        {
            // insert string from GetParameterDefinition() into sub/function/... declaration
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

            return "ByVal" + _targetDeclaration.IdentifierName + " As " + _targetDeclaration.AsTypeName;
        }
    }
}
