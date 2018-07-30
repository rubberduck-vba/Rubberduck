using System;
using Rubberduck.VBEditor;

namespace RubberduckTests
{
    public static class CodeStringExtensions
    {
        public static (string Code, Selection zPosition) RemovePseudoCaret(this string code)
        {
            var zPosition = new Selection();
            var lines = code.Split('\n');
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
            return (newCode, zPosition);
        }

        public static (string Code, Selection zPosition) InsertPseudoCaret(this string code, Selection zPosition)
        {
            var lines = code.Split('\n');
            var line = lines[zPosition.StartLine];
            lines[zPosition.StartLine] = line.Insert(zPosition.StartColumn, "|");
            return (string.Join("\n", lines), zPosition);
        }
    }
}
