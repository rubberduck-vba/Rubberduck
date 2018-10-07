using System;
using System.Linq;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.AutoComplete
{
    public abstract class AutoCompleteHandlerBase
    {
        protected AutoCompleteHandlerBase(ICodePaneHandler pane)
        {
            CodePaneHandler = pane;
        }

        protected ICodePaneHandler CodePaneHandler { get; }

        public abstract CodeString Handle(AutoCompleteEventArgs e, AutoCompleteSettings settings);
        protected bool IsInsideStringLiteral(ref CodeString currentContent)
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
    }
}