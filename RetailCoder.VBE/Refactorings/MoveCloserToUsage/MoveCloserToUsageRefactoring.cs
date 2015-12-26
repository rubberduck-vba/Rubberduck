using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageRefactoring : IRefactoring
    {
        private readonly List<Declaration> _declarations;
        private readonly IActiveCodePaneEditor _editor;
        private readonly IMessageBox _messageBox;

        public MoveCloserToUsageRefactoring(RubberduckParserState parseResult, IActiveCodePaneEditor editor, IMessageBox messageBox)
        {
            _declarations = parseResult.AllDeclarations.ToList();
            _editor = editor;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var qualifiedSelection = _editor.GetSelection();
            if (qualifiedSelection != null)
            {
                Refactor(_declarations.FindVariable(qualifiedSelection.Value));
            }
            else
            {
                _messageBox.Show("Invalid Selection.", "Rubberduck - Move Closer To Usage", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
            }
        }

        public void Refactor(QualifiedSelection selection)
        {
            Refactor(_declarations.FindVariable(selection));
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException(@"Invalid Argument", "target");
            }

            if (!target.References.Any())
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetHasNoReferences, target.IdentifierName);

                _messageBox.Show(message, RubberduckUI.MoveCloserToUsage_Caption, System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);

                return;
            }

            if (TargetIsReferencedFromMultipleMethods(target))
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetIsUsedInMultipleMethods, target.IdentifierName);
                _messageBox.Show(message, RubberduckUI.MoveCloserToUsage_Caption, System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);

                return;
            }

            MoveDeclaration(target);
        }

        private bool TargetIsReferencedFromMultipleMethods(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();

            return firstReference != null && target.References.Any(r => r.ParentScope != firstReference.ParentScope);
        }

        private void MoveDeclaration(Declaration target)
        {
            InsertDeclaration(target);
            RemoveVariable(target);
        }

        private void InsertDeclaration(Declaration target)
        {
            var firstReference = target.References.OrderBy(r => r.Selection.StartLine).First();

            var oldLines = _editor.GetLines(firstReference.Selection);
            var newLines = oldLines.Insert(firstReference.Selection.StartColumn - 1, GetDeclarationString(target));

            _editor.DeleteLines(firstReference.Selection);
            _editor.InsertLines(firstReference.Selection.StartLine, newLines);
        }

        private string GetDeclarationString(Declaration target)
        {
            return Environment.NewLine + "    Dim " + target.IdentifierName + " As " + target.AsTypeName + Environment.NewLine;
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

            if (newLines.Trim() != string.Empty)
            {
                _editor.InsertLines(selection.StartLine, newLines);
            }
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
    }
}