using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteKeyDownHandler
    {
        private readonly Func<AutoCompleteSettings> _getSettings;
        private readonly Func<List<SelfClosingPair>> _getClosingPairs;
        private readonly Func<SelfClosingPairCompletionService> _getClosingPairCompletion;
        private readonly Func<ICodeStringPrettifier> _getPrettifier;

        public AutoCompleteKeyDownHandler(Func<AutoCompleteSettings> getSettings, Func<List<SelfClosingPair>> getClosingPairs, Func<SelfClosingPairCompletionService> getClosingPairCompletion, Func<ICodeStringPrettifier> getPrettifier)
        {
            _getSettings = getSettings;
            _getClosingPairs = getClosingPairs;
            _getClosingPairCompletion = getClosingPairCompletion;
            _getPrettifier = getPrettifier;
        }

        public void Run(ICodeModule module, Selection pSelection, AutoCompleteEventArgs e)
        {
            if (!pSelection.IsSingleCharacter) { return; }

            var currentContent = module.GetLines(pSelection);
            HandleSmartConcat(e, pSelection, currentContent, module);
            if (e.Handled) { return; }

            HandleSelfClosingPairs(e, module, pSelection);
            if (e.Handled) { return; }

            //HandleSomethingElse(?)
            //if (e.Handled) { return; }
        }

        /// <summary>
        /// Adds a line continuation when {ENTER} is pressed inside a string literal.
        /// </summary>
        private void HandleSmartConcat(AutoCompleteEventArgs e, Selection pSelection, string currentContent, ICodeModule module)
        {
            var shouldHandle = _getSettings().EnableSmartConcat &&
                               e.Character == '\r' &&
                               IsInsideStringLiteral(pSelection, ref currentContent);

            var lastIndexLeftOfCaret = currentContent.Length > 2 ? currentContent.Substring(0, pSelection.StartColumn - 1).LastIndexOf('"') : 0;
            if (shouldHandle && lastIndexLeftOfCaret > 0)
            {
                var indent = currentContent.NthIndexOf('"', 1);
                var whitespace = new string(' ', indent);
                var code = $"{currentContent.Substring(0, pSelection.StartColumn - 1)}\" & _\r\n{whitespace}\"{currentContent.Substring(pSelection.StartColumn - 1)}";

                if (e.ControlDown)
                {
                    code = $"{currentContent.Substring(0, pSelection.StartColumn - 1)}\" & vbNewLine & _\r\n{whitespace}\"{currentContent.Substring(pSelection.StartColumn - 1)}";
                }

                module.ReplaceLine(pSelection.StartLine, code);
                using (var pane = module.CodePane)
                {
                    pane.Selection = new Selection(pSelection.StartLine + 1, indent + currentContent.Substring(pSelection.StartColumn - 2).Length);
                    e.Handled = true;
                }
            }
        }

        private void HandleSelfClosingPairs(AutoCompleteEventArgs e, ICodeModule module, Selection pSelection)
        {
            var original = GetEntireLogicalCodeLine(module, pSelection); // todo: see if AutoCompleteEventArgs can give us the logical line

            var prettifier = _getPrettifier();
            var scpService = _getClosingPairCompletion();

            foreach (var selfClosingPair in _getClosingPairs())
            {
                CodeString result;
                if (e.Character == '\b' && pSelection.StartColumn > 1)
                {
                    result = scpService.Execute(selfClosingPair, original, Keys.Back);
                }
                else
                {
                    result = scpService.Execute(selfClosingPair, original, e.Character);
                }

                if (!result?.Equals(default) ?? false)
                {
                    using (var pane = module.CodePane)
                    {
                        var prettified = prettifier.Prettify(module, original);
                        if (e.Character == '\b' && pSelection.StartColumn > 1)
                        {
                            result = scpService.Execute(selfClosingPair, prettified, Keys.Back);
                        }
                        else
                        {
                            result = scpService.Execute(selfClosingPair, prettified, e.Character);
                        }

                        var currentLine = result.Lines[result.CaretPosition.StartLine];
                        if (!string.IsNullOrWhiteSpace(currentLine) && currentLine.EndsWith(" ") && result.CaretPosition.StartColumn == currentLine.Length)
                        {
                            result = result.ReplaceLine(result.CaretPosition.StartLine, currentLine.TrimEnd());
                            result = new CodeString(result.Code, result.CaretPosition.ShiftLeft(), result.SnippetPosition);
                        }

                        var reprettified = prettifier.Prettify(module, result);
                        var offByOne = reprettified.Code.Length - result.Code.Length == 1;
                        if (!string.IsNullOrWhiteSpace(currentLine) && !offByOne && result.Code != reprettified.Code)
                        {
                            Debug.Assert(false, "Prettified code is off by more than one character.");
                        }

                        var finalSelection = new Selection(result.SnippetPosition.StartLine, result.CaretPosition.StartColumn + 1)
                            .ShiftRight(offByOne ? 1 : 0);
                        pane.Selection = finalSelection;
                        e.Handled = true;
                        return;
                    }
                }
            }
        }

        private CodeString GetEntireLogicalCodeLine(ICodeModule module, Selection pSelection)
        {
            var lines = new List<(int Line, string Content)>();
            var currentLineIndex = pSelection.StartLine;
            var currentLine = module.GetLines(currentLineIndex, 1);

            var caretLine = (currentLineIndex, currentLine);
            lines.Add(caretLine);

            while (currentLineIndex >= 1)
            {
                currentLineIndex--;
                if (currentLineIndex >= 1)
                {
                    currentLine = module.GetLines(currentLineIndex, 1);
                    if (currentLine.Replace("\r\n", string.Empty).EndsWith(" _"))
                    {
                        lines.Insert(0, (currentLineIndex, currentLine));
                    }
                    else
                    {
                        break;
                    }
                }
            }

            currentLineIndex = pSelection.StartLine;
            currentLine = caretLine.currentLine;
            while (currentLineIndex <= module.CountOfLines && currentLine.Replace("\r\n", string.Empty).EndsWith(" _"))
            {
                currentLineIndex++;
                if (currentLineIndex <= module.CountOfLines)
                {
                    currentLine = module.GetLines(currentLineIndex, 1);
                    lines.Add((currentLineIndex,currentLine));
                }
                else
                {
                    break;
                }
            }

            var logicalLine = string.Join("\r\n", lines.Select(e => e.Content));
            var zCaretLine = lines.IndexOf(caretLine);
            var zCaretColumn = pSelection.StartColumn - 1;

            var startLine = lines[0].Line;
            var endLine = lines[lines.Count - 1].Line;

            var result = new CodeString(logicalLine, new Selection(zCaretLine, zCaretColumn), new Selection(startLine, 1, endLine, lines[lines.Count - 1].Content.Length));
            return result;
        }

        private bool IsInsideStringLiteral(Selection pSelection, ref string currentContent)
        {
            if (!currentContent.Substring(pSelection.StartColumn - 1).Contains("\"") ||
                currentContent.StripStringLiterals().HasComment(out _))
            {
                return false;
            }

            var zSelection = pSelection.ToZeroBased();
            var leftOfCaret = currentContent.Substring(0, zSelection.StartColumn);
            var rightOfCaret = currentContent.Substring(Math.Min(zSelection.StartColumn + 1, currentContent.Length - 1));
            if (!rightOfCaret.Contains("\""))
            {
                // the string isn't terminated, but VBE would terminate it here.
                currentContent += "\"";
                rightOfCaret += "\"";
            }

            // odd number of double quotes on either side of the caret means we're inside a string literal, right?
            return (leftOfCaret.Count(c => c.Equals('"')) % 2) != 0 &&
                   (rightOfCaret.Count(c => c.Equals('"')) % 2) != 0;
        }
    }
}