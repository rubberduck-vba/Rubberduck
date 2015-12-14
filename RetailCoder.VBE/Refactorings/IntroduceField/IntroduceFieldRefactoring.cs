using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceField : IRefactoring
    {
        private readonly IList<Declaration> _declarations;
        private readonly IActiveCodePaneEditor _editor;
        private Declaration _targetDeclaration;
        private readonly IMessageBox _messageBox;

        public IntroduceField(RubberduckParserState parseResult, IActiveCodePaneEditor editor, IMessageBox messageBox)
        {
            _declarations = parseResult.AllDeclarations.ToList();
            _editor = editor;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            if (_targetDeclaration == null)
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection);
                return;
            }

            RemoveVariable();
            AddField();
        }

        public void Refactor(QualifiedSelection selection)
        {
            _targetDeclaration = FindSelection(selection);
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Invalid declaration type");
            }

            _targetDeclaration = target;
            Refactor();
        }

        private void AddField()
        {
            var module = _targetDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            module.InsertLines(module.CountOfDeclarationLines + 1, GetFieldDefinition());
        }

        private void RemoveVariable()
        {
            Selection selection;
            var declarationText = _targetDeclaration.Context.GetText();
            var multipleDeclarations = HasMultipleDeclarationsInStatement(_targetDeclaration);

            var variableStmtContext = GetVariableStmtContext(_targetDeclaration);

            if (!multipleDeclarations)
            {
                declarationText = variableStmtContext.GetText();
                selection = GetVariableStmtContextSelection(_targetDeclaration);
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
                selection = GetVariableStmtContextSelection(_targetDeclaration);
                newLines = RemoveExtraComma(_editor.GetLines(selection).Replace(oldLines, newLines));
            }

            _editor.DeleteLines(selection);
            _editor.InsertLines(selection.StartLine, newLines);
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

        private bool HasMultipleDeclarationsInStatement(Declaration target)
        {
            var statement = target.Context.Parent as VBAParser.VariableListStmtContext;

            if (statement == null) { return false; }

            return statement.children.Count(i => i is VBAParser.VariableSubStmtContext) > 1;
        }

        private string GetFieldDefinition()
        {
            if (_targetDeclaration == null) { return null; }

            return "Private " + _targetDeclaration.IdentifierName + " As " + _targetDeclaration.AsTypeName;
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
