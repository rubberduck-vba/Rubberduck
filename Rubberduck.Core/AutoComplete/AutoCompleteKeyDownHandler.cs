using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteKeyDownHandler
    {
        private static readonly IEnumerable<SelfClosingPair> SelfClosingPairs = new List<SelfClosingPair>
        {
            new SelfClosingPair('(', ')'),
            new SelfClosingPair('"', '"'),
            new SelfClosingPair('[', ']'),
            new SelfClosingPair('{', '}'),
        };

        private readonly ICodePaneHandler _codePane;
        private readonly SelfClosingPairCompletionService _scpService;

        public AutoCompleteKeyDownHandler(ICodePaneHandler codePane, SelfClosingPairCompletionService scpService)
        {
            _codePane = codePane;
            _scpService = scpService;
        }

        public AutoCompleteSettings Settings { get; internal set; }

        public void Run(AutoCompleteEventArgs e)
        {
            Selection pSelection;
            var module = e.Module;
            using (var pane = module.CodePane)
            {
                pSelection = pane.Selection;
            }

            if (pSelection.LineCount > 1)
            {
                return;
            }

            var handlers = new Action<AutoCompleteEventArgs>[]
            {
                HandleSmartConcat,
                HandleSelfClosingPairs
            };

            foreach (var handler in handlers)
            {
                handler.Invoke(e);
                if (e.Handled)
                {
                    return;
                }
            }
        }

        /// <summary>
        /// Adds a line continuation when {ENTER} is pressed inside a string literal.
        /// </summary>
        private void HandleSmartConcat(AutoCompleteEventArgs e)
        {
            if (e.Character != '\r' || (!Settings?.EnableSmartConcat ?? true))
            {
                return;
            }

            var currentContent = _codePane.GetCurrentLogicalLine(e.Module);
            var shouldHandle = IsInsideStringLiteral(ref currentContent);
            if (!shouldHandle)
            {
                return;
            }
            
            var lastIndexLeftOfCaret = currentContent.Code.Length > 2 ? currentContent.Code.Substring(0, currentContent.CaretCharIndex).LastIndexOf('"') : 0;
            if (lastIndexLeftOfCaret > 0)
            {
                var indent = currentContent.Code.NthIndexOf('"', 1);
                var whitespace = new string(' ', indent);

                var autoCode = $"\" & {(e.IsControlKeyDown ? " vbNewLine & " : string.Empty)}\" _\r\n{whitespace}\"";
                var code = $"{currentContent.Code.Substring(0, currentContent.CaretCharIndex)}{autoCode}{currentContent.Code.Substring(currentContent.CaretCharIndex + 1)}";

                var newContent = new CodeString(code, currentContent.CaretPosition, currentContent.SnippetPosition);
                _codePane.SubstituteCode(e.Module, newContent);

                var newSelection = new Selection(newContent.CaretPosition.StartLine + newContent.SnippetPosition.StartLine + 1,
                                                 newContent.Code.Substring(newContent.CaretCharIndex - 1).Length + indent);
                _codePane.SetSelection(e.Module, newSelection);
                e.Handled = true;
            }
        }

        private bool IsInsideStringLiteral(ref CodeString currentContent)
        {
            if (!currentContent.Code.Substring(currentContent.CaretPosition.StartColumn).Contains("\"") ||
                currentContent.Code.StripStringLiterals().HasComment(out _))
            {
                return false;
            }

            var leftOfCaret = currentContent.Code.Substring(0, currentContent.CaretCharIndex);
            var rightOfCaret = currentContent.Code.Substring(Math.Min(currentContent.CaretCharIndex + 1, currentContent.Code.Length - 1));
            if (!rightOfCaret.Contains("\""))
            {
                // the string isn't terminated, but VBE would terminate it here.
                currentContent = new CodeString(currentContent.Code + "\"", currentContent.CaretPosition, currentContent.SnippetPosition);
                rightOfCaret += "\"";
            }

            // odd number of double quotes on either side of the caret means we're inside a string literal, right?
            return (leftOfCaret.Count(c => c.Equals('"')) % 2) != 0 &&
                   (rightOfCaret.Count(c => c.Equals('"')) % 2) != 0;
        }

        private void HandleSelfClosingPairs(AutoCompleteEventArgs e)
        {
            var original = _codePane.GetCurrentLogicalLine(e.Module);
            foreach (var selfClosingPair in SelfClosingPairs)
            {
                CodeString result;
                if (e.Character == '\b' && original.CaretPosition.StartColumn > 1)
                {
                    result = _scpService.Execute(selfClosingPair, original, Keys.Back);
                }
                else
                {
                    result = _scpService.Execute(selfClosingPair, original, e.Character);
                }

                if (!result?.Equals(default) ?? false)
                {
                    var prettified = _codePane.Prettify(e.Module, original);
                    if (e.Character == '\b' && original.CaretPosition.StartColumn > 1)
                    {
                        result = _scpService.Execute(selfClosingPair, prettified, Keys.Back);
                    }
                    else
                    {
                        result = _scpService.Execute(selfClosingPair, prettified, e.Character);
                    }

                    var currentLine = result.Lines[result.CaretPosition.StartLine];
                    if (!string.IsNullOrWhiteSpace(currentLine) && currentLine.EndsWith(" ") &&
                        result.CaretPosition.StartColumn == currentLine.Length)
                    {
                        result = result.ReplaceLine(result.CaretPosition.StartLine, currentLine.TrimEnd());
                        result = new CodeString(result.Code, result.CaretPosition.ShiftLeft(), result.SnippetPosition);
                    }

                    var reprettified = _codePane.Prettify(e.Module, result);
                    var offByOne = reprettified.Code.Length - result.Code.Length == 1;
                    if (!string.IsNullOrWhiteSpace(currentLine) && !offByOne && result.Code != reprettified.Code)
                    {
                        Debug.Assert(false, "Prettified code is off by more than one character.");
                    }

                    var finalSelection = new Selection(result.SnippetPosition.StartLine,
                            result.CaretPosition.StartColumn + 1)
                        .ShiftRight(offByOne ? 1 : 0);
                    _codePane.SetSelection(e.Module, finalSelection);
                    e.Handled = true;
                }
            }
        }
    }
}