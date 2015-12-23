using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
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
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceField_Caption,
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
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
                newLines = RemoveExtraComma(_editor.GetLines(selection).Replace(oldLines, newLines));
            }

            _editor.DeleteLines(selection);
            _editor.InsertLines(selection.StartLine, newLines);
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

        private string GetFieldDefinition(Declaration target)
        {
            if (target == null) { return null; }

            return "Private " + target.IdentifierName + " As " + target.AsTypeName;
        }
    }
}
