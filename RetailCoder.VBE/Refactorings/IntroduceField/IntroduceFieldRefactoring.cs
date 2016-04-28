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
        private readonly VBE _vbe;
        private readonly RubberduckParserState _parserState;
        private readonly IMessageBox _messageBox;

        private IList<Declaration> _declarations;
        private Declaration _target;

        public IntroduceFieldRefactoring(VBE vbe, RubberduckParserState parserState, IMessageBox messageBox)
        {
            _vbe = vbe;
            _parserState = parserState;
            _messageBox = messageBox;
        }

        public bool CanExecute(QualifiedSelection selection)
        {
            _declarations = _parserState.AllUserDeclarations.ToList();

            _target = _declarations.FindVariable(selection);
            return _target != null;
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
            if (!CanExecute(selection))
            {
                _messageBox.Show(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            PromoteVariable(_target);
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

            _target = target;

            RemoveVariable();
            AddField();
        }

        private void AddField()
        {
            var module = _target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            module.InsertLines(module.CountOfDeclarationLines + 1, GetFieldDefinition());
        }

        private void RemoveVariable()
        {
            Selection selection;
            var declarationText = _target.Context.GetText();
            var multipleDeclarations = _target.HasMultipleDeclarationsInStatement();

            var variableStmtContext = _target.GetVariableStmtContext();

            if (!multipleDeclarations)
            {
                declarationText = variableStmtContext.GetText();
                selection = _target.GetVariableStmtContextSelection();
            }
            else
            {
                selection = new Selection(_target.Context.Start.Line, _target.Context.Start.Column,
                    _target.Context.Stop.Line, _target.Context.Stop.Column);
            }

            var oldLines = _vbe.ActiveCodePane.CodeModule.GetLines(selection);

            var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                .Remove(selection.StartColumn, declarationText.Length);

            if (multipleDeclarations)
            {
                selection = _target.GetVariableStmtContextSelection();
                newLines = RemoveExtraComma(_vbe.ActiveCodePane.CodeModule.GetLines(selection).Replace(oldLines, newLines),
                    _target.CountOfDeclarationsInStatement(), _target.IndexOfVariableDeclarationInStatement());
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

        private string GetFieldDefinition()
        {
            return "Private " + _target.IdentifierName + " As " + _target.AsTypeName;
        }
    }
}
