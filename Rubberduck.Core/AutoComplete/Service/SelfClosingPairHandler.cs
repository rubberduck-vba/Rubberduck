using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete.Service
{
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
            if (!_scpInputLookup.TryGetValue(e.Character, out var pair) && e.Character != '\b')
            {
                return false;
            }

            var original = CodePaneHandler.GetCurrentLogicalLine(e.Module);
            if (original == null || original.Lines.Length == MaximumLines)
            {
                // selection spans more than a single logical line, or
                // logical line somehow spans more than the maximum number of physical lines in a logical line of code (25).
                return false;
            }

            if (pair != null)
            {
                if (!HandleInternal(e, original, pair, out result))
                {
                    return false;
                }
            }
            else if (e.Character == '\b')
            {
                foreach (var scp in _selfClosingPairs)
                {
                    if (HandleInternal(e, original, scp, out result))
                    {
                        break;
                    }
                }

                if (result == null)
                {
                    return false;
                }
            }

            var snippetPosition = new Selection(result.SnippetPosition.StartLine, 1, result.SnippetPosition.EndLine, 1);
            result = new CodeString(result.Code, result.CaretPosition, snippetPosition);

            e.Handled = true;
            return true;
        }

        private bool HandleInternal(AutoCompleteEventArgs e, CodeString original, SelfClosingPair pair, out CodeString result)
        {
            if (!original.CaretPosition.IsSingleCharacter)
            {
                // todo: WrapSelection?
                result = null;
                return false;
            }

            var isPresent = original.CaretLine.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}");

            if (!_scpService.Execute(pair, original, e.Character, out result))
            {
                return false;
            }

            var prettified = CodePaneHandler.Prettify(e.Module, original);
            if (!isPresent && original.CaretLine.Length + 2 == prettified.CaretLine.Length &&
                prettified.CaretLine.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}"))
            {
                // prettifier just added the pair for us; likely a Sub or Function statement.
                prettified = original; // pretend this didn't happen. note: probably breaks if original has extra whitespace.
            }

            if (!_scpService.Execute(pair, prettified, e.Character, out result))
            {
                return false;
            }

            result = CodePaneHandler.Prettify(e.Module, result);

            var currentLine = result.Lines[result.CaretPosition.StartLine];
            if (!string.IsNullOrWhiteSpace(currentLine) &&
                currentLine.EndsWith(" ") &&
                result.CaretPosition.StartColumn == currentLine.Length)
            {
                result = result.ReplaceLine(result.CaretPosition.StartLine, currentLine.TrimEnd());
            }

            if (pair.OpeningChar == '(' && 
                e.Character == pair.OpeningChar &&
                !result.CaretLine.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}"))
            {
                // VBE eats it. bail out but still swallow the keypress, since we've already re-prettified.
                e.Handled = true;
                result = null;
                return false;
            }

            return true;
        }
    }
}