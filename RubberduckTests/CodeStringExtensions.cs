using System;
using Rubberduck.Common;
using Rubberduck.VBEditor;

namespace RubberduckTests
{
    public static class CodeStringExtensions
    {
        /// <summary>
        /// Creates a code string that encapsulates the caret position, indicated by a single pipe ("<c>|</c>") character.
        /// </summary>
        /// <param name="code">The code snippet string. Use a single pipe ("<c>|</c>") to indicate caret position.</param>
        /// <returns>Returns a <c>struct</c> that encapsulates a snippet of code and a cursor/caret position relative to that snippet.</returns>
        public static CodeString ToCodeString(this string code)
        {
            var zPosition = new Selection();
            var lines = (code ?? string.Empty).Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var index = line.IndexOf('|');
                if (index >= 0)
                {
                    lines[i] = line.Remove(index, 1);
                    zPosition = new Selection(i, index);
                    break;
                }
            }

            var newCode = string.Join("\n", lines);
            return new CodeString(newCode, zPosition);
        }

        public static CodeString InsertPseudoCaret(this string code, Selection zPosition)
        {
            var lines = (code ?? string.Empty).Split('\n');
            var line = lines[zPosition.StartLine];
            lines[zPosition.StartLine] = line.Insert(zPosition.StartColumn, "|");
            return new CodeString(string.Join("\n", lines), zPosition);
        }
    }
}
