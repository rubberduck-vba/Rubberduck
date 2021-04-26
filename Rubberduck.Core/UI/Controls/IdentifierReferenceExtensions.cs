using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Controls
{
    static class IdentifierReferenceExtensions
    {
        public static (string context, Selection highlight) HighlightSelection(this IdentifierReference reference, ICodeModule module)
        {
            const int maxLength = 255;
            var selection = reference.Selection;

            var lines = module.GetLines(selection.StartLine, selection.LineCount).Split('\n');

            var line = lines[0];
            var indent = line.TakeWhile(c => c.Equals(' ')).Count();

            var highlight = new Selection(
                    1, Math.Max(selection.StartColumn - indent, 1),
                    1, Math.Max(selection.EndColumn - indent, 1))
                .ToZeroBased();

            var trimmed = line.Trim();
            if (trimmed.Length > maxLength || lines.Length > 1)
            {
                trimmed = trimmed.Substring(0, maxLength) + " …";
            }
            return (trimmed, highlight);
        }

        public static (string context, Selection highlight) HighlightSelection(this ArgumentReference reference, ICodeModule module)
        {
            const int maxLength = 255;
            var selection = reference.Selection;

            var lines = module.GetLines(selection.StartLine, selection.LineCount).Split('\n');

            var line = lines[0];
            var indent = line.TakeWhile(c => c.Equals(' ')).Count();

            var highlight = new Selection(
                    1, Math.Max(selection.StartColumn - indent, 1),
                    1, Math.Max(selection.EndColumn - indent, 1))
                .ToZeroBased();

            var trimmed = line.Trim();
            if (trimmed.Length > maxLength || lines.Length > 1)
            {
                trimmed = trimmed.Substring(0, Math.Min(trimmed.Length, maxLength)) + " …";
                highlight = new Selection(1, highlight.StartColumn, 1, trimmed.Length);
            }

            if (highlight.IsSingleCharacter && highlight.StartColumn == 0)
            {
                trimmed = " " + trimmed;
                highlight = new Selection(0, 0, 0, 1);
            }
            else if (highlight.IsSingleCharacter)
            {
                highlight = new Selection(0, selection.StartColumn - 1 - indent - 1, 0, selection.StartColumn - 1 - indent);
            }
            return (trimmed, highlight);
        }
    }
}
