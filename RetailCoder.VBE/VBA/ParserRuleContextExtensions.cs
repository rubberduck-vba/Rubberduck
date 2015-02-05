using System;
using Antlr4.Runtime;
using Rubberduck.Extensions;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA
{
    public static class ParserRuleContextExtensions
    {
        public static Selection GetSelection(this ParserRuleContext context)
        {
            if (context == null)
                return Selection.Empty;

            // adding +1 because ANTLR indexes are 0-based, but VBE's are 1-based.
            return new Selection(
                context.Start.Line + 1,
                context.Start.StartIndex + 1, // todo: figure out why this is off and how to fix it
                context.Stop.Line + 1,
                context.Stop.StopIndex + 1); // todo: figure out why this is off and how to fix it
        }

        public static VBAccessibility GetAccessibility(this VisualBasic6Parser.VisibilityContext context)
        {
            if (context == null)
                return VBAccessibility.Implicit;

            return (VBAccessibility) Enum.Parse(typeof (VBAccessibility), context.GetText());
        }
    }
}