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
            foreach (var selfClosingPair in SelfClosingPairs)
            {
                var result = ExecuteSelfClosingPair(e, original, selfClosingPair);
                if (result == null)
                {
                    continue;
                }

                var prettified = CodePaneHandler.Prettify(e.Module, original);
                result = CodePaneHandler.Prettify(e.Module, ExecuteSelfClosingPair(e, prettified, selfClosingPair));
                Debug.Assert(result != null, "SCP against original was non-null, and now null against prettified.");

                var currentLine = result.Lines[result.CaretPosition.StartLine];
                if (!string.IsNullOrWhiteSpace(currentLine) && 
                    currentLine.EndsWith(" ") &&
                    result.CaretPosition.StartColumn == currentLine.Length)
                {
                    result = result.ReplaceLine(result.CaretPosition.StartLine, currentLine.TrimEnd());
                }

                e.Handled = true;
                result = new CodeString(result.Code, result.CaretPosition, new Selection(result.SnippetPosition.StartLine, 1, result.SnippetPosition.EndLine, 1));
                return result;
            }

            return null;
        }

        private CodeString ExecuteSelfClosingPair(AutoCompleteEventArgs e, CodeString original, SelfClosingPair selfClosingPair)
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

            return result;
        }
    }
}