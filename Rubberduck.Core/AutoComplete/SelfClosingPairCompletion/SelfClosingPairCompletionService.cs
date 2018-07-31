using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public class SelfClosingPairCompletionService
    {
        public (string Code, Selection CaretPosition) Execute(SelfClosingPair pair, (string, Selection) original, char input)
        {
            if (input == pair.OpeningChar)
            {
                return HandleOpeningChar(pair, original);
            }
            else if (input == pair.ClosingChar)
            {
                return HandleClosingChar(pair, original);
            }
            else
            {
                return default;
            }
        }

        public (string Code, Selection CaretPosition) Execute(SelfClosingPair pair, (string, Selection) original, Keys input)
        {
            if (input == Keys.Back)
            {
                return HandleBackspace(pair, original);
            }
            else
            {
                return default;
            }
        }

        private (string, Selection) HandleOpeningChar(SelfClosingPair pair, (string Code, Selection Position) original)
        {
            var nextPosition = original.Position.ShiftRight();
            var autoCode = new string(new[] { pair.OpeningChar, pair.ClosingChar });
            return (original.Code.Insert(original.Position.StartColumn, autoCode), nextPosition);
        }

        private (string, Selection) HandleClosingChar(SelfClosingPair pair, (string Code, Selection Position) original)
        {
            var nextPosition = original.Position.ShiftRight();
            var newCode = original.Code;

            return (newCode, nextPosition);
        }

        private (string, Selection) HandleBackspace(SelfClosingPair pair, (string Code, Selection Position) original)
        {
            var lines = original.Code.Split('\n');
            var line = lines[original.Position.StartLine];

            var previousChar = line[Math.Max(0, original.Position.StartColumn - 1)];
            var nextChar = line[Math.Min(line.Length, original.Position.StartColumn)];

            return DeleteMatchingTokens(pair, original, lines, line, previousChar, nextChar);
        }

        private static (string, Selection) DeleteMatchingTokens(SelfClosingPair pair, (string Code, Selection Position) original, string[] lines, string line, char previousChar, char nextChar)
        {
            if (previousChar == pair.OpeningChar && nextChar == pair.ClosingChar)
            {
                lines[original.Position.StartLine] = line.Remove(Math.Max(0, original.Position.StartColumn - 1), 2);
                return (string.Join("\n", lines), original.Position.ShiftLeft());
            }
            else
            {
                return default;
            }
        }

    }
}
