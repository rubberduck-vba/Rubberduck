using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceFieldRefactoring : IRefactoring
    {
        private readonly IList<Declaration> _declarations;
        private readonly VBE _vbe;
        private readonly IMessageBox _messageBox;

        public IntroduceFieldRefactoring(VBE vbe, RubberduckParserState parserState, IMessageBox messageBox)
        {
            _declarations =
                parserState.AllDeclarations.Where(i => !i.IsBuiltIn && i.DeclarationType == DeclarationType.Variable)
                    .ToList();
            _vbe = vbe;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var selection = _vbe.ActiveCodePane.GetQualifiedSelection();

            if (!selection.HasValue)
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceField_Caption,
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
            if (target.DeclarationType != DeclarationType.Variable)
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
            if (new[] { DeclarationType.ClassModule, DeclarationType.ProceduralModule }.Contains(target.ParentDeclaration.DeclarationType))
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

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

            var oldLines = _vbe.ActiveCodePane.CodeModule.GetLines(selection);

            var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                .Remove(selection.StartColumn, declarationText.Length);

            if (multipleDeclarations)
            {
                selection = target.GetVariableStmtContextSelection();
                newLines = RemoveExtraComma(_vbe.ActiveCodePane.CodeModule.GetLines(selection).Replace(oldLines, newLines),
                    target.CountOfDeclarationsInStatement(), target.IndexOfVariableDeclarationInStatement());
            }

            var newLinesWithoutExcessSpaces = newLines.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
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

        private string GetFieldDefinition(Declaration target)
        {
            return "Private " + target.IdentifierName + " As " + target.AsTypeName;
        }
    }
}
