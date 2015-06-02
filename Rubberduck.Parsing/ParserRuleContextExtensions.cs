using System;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    public static class ParserRuleContextExtensions
    {
        public static Selection GetSelection(this ParserRuleContext context)
        {
            if (context == null)
                return Selection.Home;

            // ANTLR indexes are 0-based, but VBE's are 1-based.
            // 1 is the default value that will select all lines. Replace zeroes with ones.
            // See also: https://msdn.microsoft.com/en-us/library/aa443952(v=vs.60).aspx

            var startLine = context.Start.Line == 0 ? 1 : context.Start.Line;
            var startCol = context.Start.Column + 1;
            var endLine = context.Stop.Line == 0 ? 1 : context.Stop.Line;
            var endCol = startCol + context.Stop.Text.Length;

            return new Selection(
                startLine,
                startCol,
                endLine,
                endCol
                );
        }

        public static Accessibility GetAccessibility(this VBAParser.VisibilityContext context)
        {
            if (context == null)
                return Accessibility.Implicit;

            return (Accessibility) Enum.Parse(typeof (Accessibility), context.GetText());
        }

        public static string Signature(this VBAParser.FunctionStmtContext context)
        {
            var visibility = context.visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.ambiguousIdentifier().GetText();
            var argsText = context.argList().GetText();
            
            var asType = context.asTypeClause();
            var asTypeText = asType == null ? string.Empty : asType.GetText();

            return (visibilityText + ' ' + Tokens.Function + ' ' + identifierText + argsText + ' ' + asTypeText).Trim();
        }

        public static string Signature(this VBAParser.SubStmtContext context)
        {
            var visibility = context.visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.ambiguousIdentifier().GetText();
            var argsText = context.argList().GetText();

            return (visibilityText + ' ' + Tokens.Sub + ' ' + identifierText + argsText).Trim();
        }

        public static string Signature(this VBAParser.PropertyGetStmtContext context)
        {
            var visibility = context.visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.ambiguousIdentifier().GetText();
            var argsText = context.argList().GetText();

            var asType = context.asTypeClause();
            var asTypeText = asType == null ? string.Empty : asType.GetText();

            return (visibilityText + ' ' + Tokens.Property + ' ' + Tokens.Get + ' ' + identifierText + argsText + ' ' + asTypeText).Trim();
        }

        public static string Signature(this VBAParser.PropertyLetStmtContext context)
        {
            var visibility = context.visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.ambiguousIdentifier().GetText();
            var argsText = context.argList().GetText();

            return (visibilityText + ' ' + Tokens.Property + ' ' + Tokens.Let + ' ' + identifierText + argsText).Trim();
        }

        public static string Signature(this VBAParser.PropertySetStmtContext context)
        {
            var visibility = context.visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.ambiguousIdentifier().GetText();
            var argsText = context.argList().GetText();

            return (visibilityText + ' ' + Tokens.Property + ' ' + Tokens.Set + ' ' + identifierText + argsText).Trim();
        }
    }
}