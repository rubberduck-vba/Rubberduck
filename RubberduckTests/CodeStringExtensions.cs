using Rubberduck.Common;
using Rubberduck.VBEditor;

namespace RubberduckTests
{
    public static class CodeStringExtensions
    {
        /// <summary>
        /// Creates a code string that encapsulates the caret position, indicated by the <see cref="TestCodeString.PseudoCaret"/> character.
        /// </summary>
        /// <param name="code">The code snippet string. Use <see cref="TestCodeString.PseudoCaret"/> to indicate caret position.</param>
        /// <returns>Returns a <c>struct</c> that encapsulates a snippet of code and a cursor/caret position relative to that snippet.</returns>
        public static TestCodeString ToCodeString(this string code)
        {
            var zPosition = new Selection();
            var lines = (code ?? string.Empty).Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var index = line.IndexOf(TestCodeString.PseudoCaret);
                if (index >= 0)
                {
                    lines[i] = line.Remove(index, 1);
                    zPosition = new Selection(i, index);
                    break;
                }
            }

            var newCode = string.Join("\n", lines);
            return new TestCodeString(newCode, zPosition);
        }
    }
}
