using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageRefactoring : IRefactoring
    {
        private readonly List<Declaration> _declarations;
        private readonly VBE _vbe;
        private readonly RubberduckParserState _parseResult;
        private readonly IMessageBox _messageBox;

        public MoveCloserToUsageRefactoring(VBE vbe, RubberduckParserState parseResult, IMessageBox messageBox)
        {
            _declarations = parseResult.AllUserDeclarations.ToList();
            _vbe = vbe;
            _parseResult = parseResult;
            _messageBox = messageBox;
        }

        public bool CanExecute(QualifiedSelection selection)
        {
            return false;
        }

        public void Refactor()
        {
            var qualifiedSelection = _vbe.ActiveCodePane.CodeModule.GetSelection();
            if (qualifiedSelection != null)
            {
                Refactor(_declarations.FindVariable(qualifiedSelection.Value));
            }
            else
            {
                _messageBox.Show(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.MoveCloserToUsage_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }
        }

        public void Refactor(QualifiedSelection selection)
        {
            var target = _declarations.FindVariable(selection);

            if (target == null)
            {
                _messageBox.Show(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            MoveCloserToUsage(target);
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                _messageBox.Show(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.IntroduceParameter_Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                // ReSharper disable once LocalizableElement
                throw new ArgumentException("Invalid Argument. DeclarationType must be 'Variable'", "target");
            }

            MoveCloserToUsage(target);
        }

        private bool TargetIsReferencedFromMultipleMethods(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();

            return firstReference != null && target.References.Any(r => r.ParentScoping != firstReference.ParentScoping);
        }

        private void MoveCloserToUsage(Declaration target)
        {
            if (!target.References.Any())
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetHasNoReferences, target.IdentifierName);

                _messageBox.Show(message, RubberduckUI.MoveCloserToUsage_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);

                return;
            }

            if (TargetIsReferencedFromMultipleMethods(target))
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetIsUsedInMultipleMethods, target.IdentifierName);
                _messageBox.Show(message, RubberduckUI.MoveCloserToUsage_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);

                return;
            }

            // it doesn't make sense to do it backwards, but we need to work from the bottom up so our selections are accurate
            InsertDeclaration(target);

            if (target.QualifiedName.QualifiedModuleName.Component !=
                target.References.First().QualifiedModuleName.Component)
            {
                _parseResult.StateChanged += (sender, e) =>
                {
                    if (e.State == ParserState.Ready)
                    {
                        foreach (var newTarget in _parseResult.AllUserDeclarations.Where(newTarget => newTarget.ComponentName == target.ComponentName &&
                                                                                                      newTarget.IdentifierName == target.IdentifierName &&
                                                                                                      newTarget.ParentScope == target.ParentScope &&
                                                                                                      newTarget.Project == target.Project &&
                                                                                                      Equals(newTarget.Selection, target.Selection)))
                        {
                            UpdateCallsToOtherModule(newTarget.References);
                            RemoveField(newTarget);
                            break;
                        }
                    }
                };

                _parseResult.OnParseRequested(this);
            }
            else
            {
                RemoveField(target);
            }
        }

        private void InsertDeclaration(Declaration target)
        {
            var module = target.References.First().QualifiedModuleName.Component.CodeModule;

            var firstReference = target.References.OrderBy(r => r.Selection.StartLine).First();
            var beginningOfInstructionSelection = GetBeginningOfInstructionSelection(firstReference);

            var oldLines = module.Lines[beginningOfInstructionSelection.StartLine, beginningOfInstructionSelection.LineCount];
            var newLines = oldLines.Insert(beginningOfInstructionSelection.StartColumn - 1, GetDeclarationString(target));

            var newLinesWithoutStringLiterals = newLines.StripStringLiterals();

            var lastIndexOfColon = newLinesWithoutStringLiterals.LastIndexOf(':');
            while (lastIndexOfColon != -1)
            {
                var numberOfCharsToRemove = lastIndexOfColon == newLines.Length - 1 || newLines[lastIndexOfColon + 1] != ' '
                    ? 1
                    : 2;

                newLinesWithoutStringLiterals = newLinesWithoutStringLiterals
                        .Remove(lastIndexOfColon, numberOfCharsToRemove)
                        .Insert(lastIndexOfColon, Environment.NewLine);

                newLines = newLines
                        .Remove(lastIndexOfColon, numberOfCharsToRemove)
                        .Insert(lastIndexOfColon, Environment.NewLine);

                lastIndexOfColon = newLinesWithoutStringLiterals.LastIndexOf(':');
            }

            module.DeleteLines(beginningOfInstructionSelection.StartLine, beginningOfInstructionSelection.LineCount);
            module.InsertLines(beginningOfInstructionSelection.StartLine, newLines);
        }

        private Selection GetBeginningOfInstructionSelection(IdentifierReference reference)
        {
            var referenceSelection = reference.Selection;
            var module = reference.QualifiedModuleName.Component.CodeModule;

            var currentLine = referenceSelection.StartLine;

            var codeLine = module.Lines[currentLine, 1].StripStringLiterals();
            while (codeLine.Remove(referenceSelection.StartColumn).LastIndexOf(':') == -1)
            {
                codeLine = module.Lines[--currentLine, 1].StripStringLiterals();
                if (!codeLine.EndsWith(" _"))
                {
                    return new Selection(currentLine + 1, 1, currentLine + 1, 1);
                }
            }

            var index = codeLine.Remove(referenceSelection.StartColumn).LastIndexOf(':') + 1;
            return new Selection(currentLine, index, currentLine, index);
        }

        private string GetDeclarationString(Declaration target)
        {
            return Environment.NewLine + "    Dim " + target.IdentifierName + " As " + target.AsTypeName + Environment.NewLine;
        }

        private void RemoveField(Declaration target)
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

            var module = target.QualifiedName.QualifiedModuleName.Component.CodeModule;

            var oldLines = module.Lines[selection.StartLine, selection.LineCount];

            var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                .Remove(selection.StartColumn, declarationText.Length);

            if (multipleDeclarations)
            {
                selection = target.GetVariableStmtContextSelection();
                newLines = RemoveExtraComma(module.Lines[selection.StartLine, selection.LineCount].Replace(oldLines, newLines),
                    target.CountOfDeclarationsInStatement(), target.IndexOfVariableDeclarationInStatement());
            }

            var adjustedLines =
                newLines.Split(new[] {Environment.NewLine}, StringSplitOptions.None)
                    .Select(s => s.EndsWith(" _") ? s.Remove(s.Length - 2) : s)
                    .Where(s => s.Trim() != string.Empty)
                    .ToList();

            module.DeleteLines(selection.StartLine, selection.LineCount);

            if (adjustedLines.Any())
            {
                module.InsertLines(selection.StartLine, string.Join(string.Empty, adjustedLines));
            }
        }

        private string RemoveExtraComma(string str, int numParams, int indexRemoved)
        {
            /* Example use cases for this method (fields and variables):
             * Dim fizz as Boolean, dizz as Double
             * Private fizz as Boolean, dizz as Double
             * Public fizz as Boolean, _
             *        dizz as Double
             * Private fizz as Boolean _
             *         , dizz as Double _
             *         , iizz as Integer

             * Before this method is called, the parameter to be removed has 
             * already been removed.  This means 'str' will look like:
             * Dim fizz as Boolean, 
             * Private , dizz as Double
             * Public fizz as Boolean, _
             *        
             * Private  _
             *         , dizz as Double _
             *         , iizz as Integer

             * This method is responsible for removing the redundant comma
             * and returning a string similar to:
             * Dim fizz as Boolean
             * Private dizz as Double
             * Public fizz as Boolean _
             *        
             * Private  _
             *          dizz as Double _
             *         , iizz as Integer
             */

            var commaToRemove = numParams == indexRemoved ? indexRemoved - 1 : indexRemoved;

            return str.Remove(str.NthIndexOf(',', commaToRemove), 1);
        }

        private void UpdateCallsToOtherModule(IEnumerable <IdentifierReference> references)
        {
            var identifierReferences = references.ToList();

            var module = identifierReferences[0].QualifiedModuleName.Component.CodeModule;

            foreach (var reference in identifierReferences.OrderByDescending(o => o.Selection.StartLine).ThenByDescending(t => t.Selection.StartColumn))
            {
                var parent = reference.Context.Parent;
                while (!(parent is VBAParser.ICS_S_MembersCallContext))
                {
                    parent = parent.Parent;
                }

                var parentSelection = ((VBAParser.ICS_S_MembersCallContext)parent).GetSelection();
                
                var oldText = module.Lines[parentSelection.StartLine, parentSelection.LineCount];
                string newText;

                if (parentSelection.LineCount == 1)
                {
                    newText = oldText.Remove(parentSelection.StartColumn - 1,
                        parentSelection.EndColumn - parentSelection.StartColumn);
                }
                else
                {
                    var lines = oldText.Split(new[] {" _" + Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);

                    newText = lines.First().Remove(parentSelection.StartColumn - 1);
                    newText += lines.Last().Remove(0, parentSelection.EndColumn - 1);
                }

                newText = newText.Insert(parentSelection.StartColumn - 1, reference.IdentifierName);

                module.DeleteLines(parentSelection.StartLine, parentSelection.LineCount);
                module.InsertLines(parentSelection.StartLine, newText);
            }
        }
    }
}