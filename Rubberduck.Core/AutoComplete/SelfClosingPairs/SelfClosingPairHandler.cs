using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete.SelfClosingPairs
{
    /// <summary>
    /// An AC handler that automatically closes certain specific "pairs" of characters, e.g. double quotes, or parentheses.
    /// </summary>
    public class SelfClosingPairHandler : AutoCompleteHandlerBase
    {
        private const int MaximumLines = 25;

        private readonly IReadOnlyList<SelfClosingPair> _selfClosingPairs;
        private readonly IDictionary<char, SelfClosingPair> _scpInputLookup;
        private readonly SelfClosingPairCompletionService _scpService;

        public SelfClosingPairHandler(ICodePaneHandler pane, SelfClosingPairCompletionService scpService)
            : base(pane)
        {
            _selfClosingPairs = new[]
            {
                new SelfClosingPair('(', ')'),
                new SelfClosingPair('"', '"'),
                new SelfClosingPair('[', ']'),
                new SelfClosingPair('{', '}'),
            };
            _scpInputLookup = _selfClosingPairs
                .Select(p => new {Key = p.OpeningChar, Pair = p})
                .Union(_selfClosingPairs.Where(p => !p.IsSymetric).Select(p => new {Key = p.ClosingChar, Pair = p}))
                .ToDictionary(p => p.Key, p => p.Pair);

            _scpService = scpService;
        }

        public override bool Handle(AutoCompleteEventArgs e, AutoCompleteSettings settings, out CodeString result)
        {
            result = null;
            if (!settings.SelfClosingPairs.IsEnabled || !_scpInputLookup.TryGetValue(e.Character, out var pair) && e.Character != '\b')
            {
                // not an interesting keypress.
                return false;
            }

            var original = CodePaneHandler.GetCurrentLogicalLine(e.Module);
            if (original == null || original.Lines.Length == MaximumLines)
            {
                // selection spans more than a single logical line, or
                // logical line somehow spans more than the maximum number of physical lines in a logical line of code (25).
                return false;
            }

            if (!original.CaretPosition.IsSingleCharacter)
            {
                // here would be an opportunity to "wrap selection" with a SCP.
                // todo: WrapSelectionWith(pair)?
                result = null;
                return false;
            }

            if (pair != null)
            {
                // found a SCP for the input key; see if we should handle it:
                if (!HandleInternal(e, original, pair, out result))
                {
                    return false;
                }
            }
            else if (e.Character == '\b')
            {
                // backspace - see if SCP logic needs to intervene:
                foreach (var scp in _selfClosingPairs)
                {
                    if (HandleInternal(e, original, scp, out result))
                    {
                        break;
                    }
                }
            }

            if (result == null)
            {
                // no meaningful output; let the input be handled by another handler, maybe.
                return false;
            }

            // 1-based selection span in the code pane starts at column 1 but really encompasses the entire line.
            var snippetPosition = new Selection(result.SnippetPosition.StartLine, 1, result.SnippetPosition.EndLine, 1);
            result = new CodeString(result.Code, result.CaretPosition, snippetPosition);
            _scpService.ShowQuickInfo();
            e.Handled = true;
            return true;
        }

        private bool HandleInternal(AutoCompleteEventArgs e, CodeString original, SelfClosingPair pair, out CodeString result)
        {
            // if executing the SCP against the original code yields no result, we need to bail out.
            if (!_scpService.Execute(pair, original, e.Character, out result))
            {
                return false;
            }

            // let the VBE alter the original code if it wants to, then work with the prettified code.
            var prettified = CodePaneHandler.Prettify(e.Module, original);

            var isPresent = original.CaretLine.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}");
            if (!isPresent && original.CaretLine.Length + 2 == prettified.CaretLine.Length &&
                prettified.CaretLine.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}"))
            {
                // prettifier just added the pair for us; likely a Sub or Function statement.
                prettified = original; // pretend this didn't happen; we need to work out the caret position anyway.
            }

            if (prettified.CaretLine.Length == 0)
            {
                // prettifier destroyed the indent. need to reinstate it now.
                prettified = prettified.ReplaceLine(
                    index:prettified.CaretPosition.StartLine,
                    content:new string(' ', original.CaretLine.TakeWhile(c => c == ' ').Count())
                );
            }

            if (original.CaretLine.EndsWith(" ") && 
                string.Equals(original.CaretLine, prettified.CaretLine + " ", StringComparison.InvariantCultureIgnoreCase))
            {
                prettified = original;
            }

            // if executing the SCP against the prettified code yields no result, we need to bail out.
            if (!_scpService.Execute(pair, prettified, e.Character, out result))
            {
                return false;
            }

            var reprettified = CodePaneHandler.Prettify(e.Module, result);
            if (pair.OpeningChar == '(' && e.Character == pair.OpeningChar)
            {
                if (string.Equals(reprettified.Code, result.Code, StringComparison.InvariantCultureIgnoreCase))
                {
                    e.Handled = true;
                    result = reprettified;
                    return true;
                }

                // VBE eats it. bail out but don't swallow the keypress.
                e.Handled = false;
                result = null;
                return false;
            }

            var currentLine = reprettified.Lines[reprettified.CaretPosition.StartLine];
            if (!string.IsNullOrWhiteSpace(currentLine) &&
                currentLine.EndsWith(" ") &&
                reprettified.CaretPosition.StartColumn == currentLine.Length)
            {
                result = reprettified.ReplaceLine(reprettified.CaretPosition.StartLine, currentLine.TrimEnd());
            }

            if (pair.OpeningChar == '(' && 
                e.Character == pair.OpeningChar &&
                !result.CaretLine.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}"))
            {
                // VBE eats it. bail out but still swallow the keypress; we already prettified the opening character into the editor.
                e.Handled = true;
                result = null;
                return false;
            }

            return true;
        }
    }
}