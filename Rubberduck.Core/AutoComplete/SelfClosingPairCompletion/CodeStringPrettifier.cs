using System;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public class CodeStringPrettifier : ICodeStringPrettifier
    {
        public CodeString Prettify(ICodeModule module, CodeString original)
        {
            var originalCode = original.Code.Replace("\r", string.Empty).Split('\n');
            var originalPosition = original.CaretPosition.StartColumn;
            var originalNonWhitespaceCharacters = 0;
            for (var i = 0; i <= originalPosition - 1; i++)
            {
                if (originalCode[original.CaretPosition.StartLine][i] != ' ')
                {
                    originalNonWhitespaceCharacters++;
                }
            }

            var indent = originalCode[original.CaretPosition.StartLine].TakeWhile(c => c == ' ').Count();
            
            module.DeleteLines(original.SnippetPosition.StartLine, original.SnippetPosition.LineCount);
            module.InsertLines(original.SnippetPosition.StartLine, string.Join("\r\n", originalCode));
            var prettifiedCode = module.GetLines(original.SnippetPosition).Replace("\r", string.Empty).Split('\n');

            var prettifiedNonWhitespaceCharacters = 0;
            var prettifiedCaretCharIndex = 0;
            for (var i = 0; i < prettifiedCode[original.CaretPosition.StartLine].Length; i++)
            {
                if (prettifiedCode[original.CaretPosition.StartLine][i] != ' ')
                {
                    prettifiedNonWhitespaceCharacters++;
                    if (prettifiedNonWhitespaceCharacters == originalNonWhitespaceCharacters)
                    {
                        prettifiedCaretCharIndex = i;
                        break;
                    }
                }
            }

            var prettifiedPosition = new Selection(
                original.SnippetPosition.StartLine - 1 + original.CaretPosition.StartLine, 
                prettifiedCode[original.CaretPosition.StartLine].Trim().Length == 0 ? indent : Math.Min(prettifiedCode[original.CaretPosition.StartLine].Length, prettifiedCaretCharIndex + 1))
                .ToOneBased();

            using (var pane = module.CodePane)
            {
                pane.Selection = prettifiedPosition;
            }

            var caretPosition = new Selection(original.CaretPosition.StartLine, prettifiedPosition.StartColumn - 1);
            var snippetPosition = new Selection(original.SnippetPosition.StartLine, original.SnippetPosition.StartColumn, original.SnippetPosition.EndLine, prettifiedCode[prettifiedCode.Length - 1].Length);
            var result = new CodeString(string.Join("\r\n", prettifiedCode), caretPosition, snippetPosition);
            return result;
        }
    }
}
