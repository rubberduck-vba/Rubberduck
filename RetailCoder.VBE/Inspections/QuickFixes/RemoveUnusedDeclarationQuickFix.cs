using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// A code inspection quickfix that removes an unused identifier declaration.
    /// </summary>
    public class RemoveUnusedDeclarationQuickFix : QuickFixBase
    {
        private readonly Declaration _target;

        public RemoveUnusedDeclarationQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target)
            : base(context, selection, InspectionsUI.RemoveUnusedDeclarationQuickFix)
        {
            _target = target;
        }

        public override void Fix()
        {
            if (_target.DeclarationType == DeclarationType.Variable || _target.DeclarationType == DeclarationType.Constant)
            {
                RemoveVariable(_target);
            }
            else
            {
                var module = Selection.QualifiedName.Component.CodeModule;
                {
                    var selection = Selection.Selection;
                    var originalCodeLines = module.GetLines(selection.StartLine, selection.LineCount);

                    var originalInstruction = Context.GetText();
                    module.DeleteLines(selection.StartLine, selection.LineCount);

                    var newCodeLines = originalCodeLines.Replace(originalInstruction, string.Empty);
                    if (!string.IsNullOrEmpty(newCodeLines))
                    {
                        module.InsertLines(selection.StartLine, newCodeLines);
                    }
                }
            }
        }

        private void RemoveVariable(Declaration target)
        {
            Selection selection;
            var declarationText = target.Context.GetText().Replace(" _" + Environment.NewLine, string.Empty);
            var multipleDeclarations = target.DeclarationType == DeclarationType.Variable && target.HasMultipleDeclarationsInStatement();

            if (!multipleDeclarations)
            {
                declarationText = GetStmtContext(target).GetText().Replace(" _" + Environment.NewLine, string.Empty);
                selection = GetStmtContextSelection(target);
            }
            else
            {
                selection = new Selection(target.Context.Start.Line, target.Context.Start.Column,
                    target.Context.Stop.Line, target.Context.Stop.Column);
            }

            var module = target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            {
                var oldLines = module.GetLines(selection);

                var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                    .Remove(selection.StartColumn, declarationText.Length);

                if (multipleDeclarations)
                {
                    selection = GetStmtContextSelection(target);
                    newLines = RemoveExtraComma(module.GetLines(selection).Replace(oldLines, newLines),
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

                // remove all lines with only whitespace
                newLinesWithoutExcessSpaces = newLinesWithoutExcessSpaces.Where(str => str.Any(c => !char.IsWhiteSpace(c))).ToArray();

                module.DeleteLines(selection);
                if (newLinesWithoutExcessSpaces.Any())
                {
                    module.InsertLines(selection.StartLine, string.Join(Environment.NewLine, newLinesWithoutExcessSpaces));
                }
            }
        }

        private Selection GetStmtContextSelection(Declaration target)
        {
            return target.DeclarationType == DeclarationType.Variable
                ? target.GetVariableStmtContextSelection()
                : target.GetConstStmtContextSelection();
        }

        private ParserRuleContext GetStmtContext(Declaration target)
        {
            return target.DeclarationType == DeclarationType.Variable
                ? (ParserRuleContext)target.GetVariableStmtContext()
                : (ParserRuleContext)target.GetConstStmtContext();
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