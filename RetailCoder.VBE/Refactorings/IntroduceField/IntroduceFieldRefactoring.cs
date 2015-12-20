using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
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
        private readonly IMessageBox _messageBox;

        public IntroduceField(RubberduckParserState parserState, IActiveCodePaneEditor editor, IMessageBox messageBox)
        {
            _declarations =
                parserState.AllDeclarations.Where(i => !i.IsBuiltIn && i.DeclarationType == DeclarationType.Variable)
                    .ToList();
            _editor = editor;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var selection = _editor.GetSelection();
            
            if (!selection.HasValue)
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceField_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
            AddField(target);
        }

        private void AddField(Declaration target)
        {
            var module = target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            module.InsertLines(module.CountOfDeclarationLines + 1, GetFieldDefinition(target));
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
                throw new NullReferenceException("Statement not found");
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

            return statement != null && statement.children.Count(i => i is VBAParser.VariableSubStmtContext) > 1;
        }

        private string GetFieldDefinition(Declaration target)
        {
            if (target == null) { return null; }

            return "Private " + target.IdentifierName + " As " + target.AsTypeName;
        }

        private Declaration FindSelection(QualifiedSelection selection)
        {
            var target = _declarations
                .FirstOrDefault(item => item.IsSelected(selection) || item.References.Any(r => r.IsSelected(selection)));

            if (target != null) { return target; }

            var targets = _declarations.Where(item => item.ComponentName == selection.QualifiedName.ComponentName);

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
