using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete.Service
{
    public class SelfClosingPairHandler : AutoCompleteHandlerBase
    {
        private static readonly IEnumerable<SelfClosingPair> SelfClosingPairs = new List<SelfClosingPair>
        {
            new SelfClosingPair('(', ')'),
            new SelfClosingPair('"', '"'),
            new SelfClosingPair('[', ']'),
            new SelfClosingPair('{', '}'),
        };

        private readonly SelfClosingPairCompletionService _scpService;

        public SelfClosingPairHandler(ICodePaneHandler pane, SelfClosingPairCompletionService scpService)
            : base(pane)
        {
            _scpService = scpService;
        }

        public override CodeString Handle(AutoCompleteEventArgs e, AutoCompleteSettings settings)
        {
            var original = CodePaneHandler.GetCurrentLogicalLine(e.Module);
            foreach (var pair in SelfClosingPairs)
            {
                var isPresent = original.CaretLine.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}");

                var result = ExecuteSelfClosingPair(e, original, pair);
                if (result == null)
                {
                    continue;
                }

                var prettified = CodePaneHandler.Prettify(e.Module, original);
                if (!isPresent && original.CaretLine.Length + 2 == prettified.CaretLine.Length && 
                    prettified.CaretLine.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}"))
                {
                    // prettifier just added the pair for us; likely a Sub or Function statement.
                    prettified = original; // pretend this didn't happen. note: probably breaks if original has extra whitespace.
                }

                result = ExecuteSelfClosingPair(e, prettified, pair);
                if (result == null)
                {
                    continue;
                }

                result = CodePaneHandler.Prettify(e.Module, result);

                var currentLine = result.Lines[result.CaretPosition.StartLine];
                if (!string.IsNullOrWhiteSpace(currentLine) && 
                    currentLine.EndsWith(" ") &&
                    result.CaretPosition.StartColumn == currentLine.Length)
                {
                    result = result.ReplaceLine(result.CaretPosition.StartLine, currentLine.TrimEnd());
                }

                if (pair.OpeningChar == '(' && e.Character != '\b' && !result.CaretLine.EndsWith($"{pair.OpeningChar}{pair.ClosingChar}"))
                {
                    // VBE eats it. just bail out.
                    return null;
                }

                e.Handled = true;
                result = new CodeString(result.Code, result.CaretPosition, new Selection(result.SnippetPosition.StartLine, 1, result.SnippetPosition.EndLine, 1));
                return result;
            }

            return null;
        }

        private CodeString ExecuteSelfClosingPair(AutoCompleteEventArgs e, CodeString original, SelfClosingPair pair)
        {
            CodeString result;
            if (e.Character == '\b' && original.CaretPosition.StartColumn > 1)
            {
                result = _scpService.Execute(pair, original, Keys.Back);
            }
            else
            {
                result = _scpService.Execute(pair, original, e.Character);
            }

            return result;
        }
    }
}