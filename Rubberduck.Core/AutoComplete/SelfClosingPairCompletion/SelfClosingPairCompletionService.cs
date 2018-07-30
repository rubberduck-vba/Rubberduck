using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public class SelfClosingPairCompletionService
    {
        public (string, Selection) Execute(SelfClosingPair pair, (string, Selection) original, char input)
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
                return original;
            }
        }

        private (string, Selection) HandleOpeningChar(SelfClosingPair pair, (string Code, Selection Position) original)
        {
            var nextPosition = original.Position.ShiftRight();

            return (original.Code.Insert(original.Position.StartColumn, pair.OpeningChar.ToString()), nextPosition);
        }

        private (string, Selection) HandleClosingChar(SelfClosingPair pair, (string Code, Selection Position) original)
        {
            var nextPosition = original.Position.ShiftRight();
            var newCode = original.Code;

            return (newCode, nextPosition);
        }
    }
}
