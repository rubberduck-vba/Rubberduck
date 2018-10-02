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
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteKeyDownHandler
    {
        private readonly ICodePaneHandler _codePane;
        private readonly Func<AutoCompleteSettings> _getSettings;
        private readonly Func<List<SelfClosingPair>> _getClosingPairs;
        private readonly Func<SelfClosingPairCompletionService> _getClosingPairCompletion;

        public AutoCompleteKeyDownHandler(ICodePaneHandler codePane, Func<AutoCompleteSettings> getSettings,
            Func<List<SelfClosingPair>> getClosingPairs,
            Func<SelfClosingPairCompletionService> getClosingPairCompletion)
        {
            _codePane = codePane;
            _getSettings = getSettings;
            _getClosingPairs = getClosingPairs;
            _getClosingPairCompletion = getClosingPairCompletion;
        }

        public void Run(QualifiedModuleName module, Selection pSelection, AutoCompleteEventArgs e)
        {
            if (!pSelection.IsSingleCharacter)
            {
                return;
            }

            HandleSmartConcat(e, module);
            if (e.Handled)
            {
                return;
            }

            HandleSelfClosingPairs(e, module);
            if (e.Handled)
            {
                // redundant now, bug if deleted and not reinstated later.
                return;
            }
        }

        /// <summary>
        /// Adds a line continuation when {ENTER} is pressed inside a string literal.
        /// </summary>
        private void HandleSmartConcat(AutoCompleteEventArgs e, QualifiedModuleName module)
        {
            var currentContent = _codePane.GetCurrentLogicalLine(module);
            var shouldHandle = _getSettings().EnableSmartConcat &&
                               e.Character == '\r' &&
                               IsInsideStringLiteral(ref currentContent);
            
            var lastIndexLeftOfCaret = currentContent.Length > 2 ? currentContent.Substring(0, pSelection.StartColumn - 1).LastIndexOf('"') : 0;
            if (shouldHandle && lastIndexLeftOfCaret > 0)
            {
                var indent = currentContent.NthIndexOf('"', 1);
                var whitespace = new string(' ', indent);
                var code =
                    $"{currentContent.Substring(0, pSelection.StartColumn - 1)}\" & _\r\n{whitespace}\"{currentContent.Substring(pSelection.StartColumn - 1)}";

                if (e.ControlDown)
                {
                    code =
                        $"{currentContent.Substring(0, pSelection.StartColumn - 1)}\" & vbNewLine & _\r\n{whitespace}\"{currentContent.Substring(pSelection.StartColumn - 1)}";
                }

                module.ReplaceLine(pSelection.StartLine, code);
                using (var pane = module.CodePane)
                {
                    pane.Selection = new Selection(pSelection.StartLine + 1,
                        indent + currentContent.Substring(pSelection.StartColumn - 2).Length);
                    e.Handled = true;
                }
            }
        }

        private bool IsInsideStringLiteral(ref CodeString logicalLine)
        {
            var caretCharIndex = logicalLine.CaretCharIndex;
            if (!logicalLine.Code.Substring(caretCharIndex).Contains("\"") ||
                logicalLine.Code.StripStringLiterals().HasComment(out _))
            {
                return false;
            }

            var leftOfCaret = logicalLine.Code.Substring(0, caretCharIndex);
            var rightOfCaret = logicalLine.Code.Substring(Math.Min(caretCharIndex + 1, logicalLine.Code.Length - 1));
            if (!rightOfCaret.Contains("\""))
            {
                // the string isn't terminated, but VBE would terminate it here.
                logicalLine += "\"";
                rightOfCaret += "\"";
            }

            // odd number of double quotes on either side of the caret means we're inside a string literal, right?
            return (leftOfCaret.Count(c => c.Equals('"')) % 2) != 0 &&
                   (rightOfCaret.Count(c => c.Equals('"')) % 2) != 0;
        }

        private void HandleSelfClosingPairs(AutoCompleteEventArgs e, QualifiedModuleName module)
        {
            var original = _codePane.GetCurrentLogicalLine(module);
            var scpService = _getClosingPairCompletion();

            foreach (var selfClosingPair in _getClosingPairs())
            {
                CodeString result;
                if (e.Character == '\b' && original.CaretPosition.StartColumn > 1)
                {
                    result = scpService.Execute(selfClosingPair, original, Keys.Back);
                }
                else
                {
                    result = scpService.Execute(selfClosingPair, original, e.Character);
                }

                if (!result?.Equals(default) ?? false)
                {
                    var prettified = _codePane.Prettify(module, original);
                    if (e.Character == '\b' && original.CaretPosition.StartColumn > 1)
                    {
                        result = scpService.Execute(selfClosingPair, prettified, Keys.Back);
                    }
                    else
                    {
                        result = scpService.Execute(selfClosingPair, prettified, e.Character);
                    }

                    var currentLine = result.Lines[result.CaretPosition.StartLine];
                    if (!string.IsNullOrWhiteSpace(currentLine) && currentLine.EndsWith(" ") &&
                        result.CaretPosition.StartColumn == currentLine.Length)
                    {
                        result = result.ReplaceLine(result.CaretPosition.StartLine, currentLine.TrimEnd());
                        result = new CodeString(result.Code, result.CaretPosition.ShiftLeft(), result.SnippetPosition);
                    }

                    var reprettified = _codePane.Prettify(module, result);
                    var offByOne = reprettified.Code.Length - result.Code.Length == 1;
                    if (!string.IsNullOrWhiteSpace(currentLine) && !offByOne && result.Code != reprettified.Code)
                    {
                        Debug.Assert(false, "Prettified code is off by more than one character.");
                    }

                    var finalSelection = new Selection(result.SnippetPosition.StartLine,
                            result.CaretPosition.StartColumn + 1)
                        .ShiftRight(offByOne ? 1 : 0);
                    _codePane.SetSelection(module, finalSelection);
                    e.Handled = true;
                    return;
                }
            }
        }
    }
}