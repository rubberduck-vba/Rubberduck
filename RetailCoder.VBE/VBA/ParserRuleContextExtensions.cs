using System;
using Antlr4.Runtime;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA
{
    public static class ParserRuleContextExtensions
    {
        public static QualifiedContext<TContext> ToQualifiedContext<TContext>(this TContext context, QualifiedModuleName name)
        {
            return new QualifiedContext<TContext>(name, context);
        }

        public static QualifiedSelection GetQualifiedSelection(this ParserRuleContext context, QualifiedModuleName name)
        {
            var selection = context.GetSelection();
            return new QualifiedSelection(name, selection);
        }

        public static Selection GetSelection(this ParserRuleContext context)
        {
            if (context == null)
                return Selection.Empty;

            // ANTLR indexes are 0-based, but VBE's are 1-based.
            // 1 is the default value that will select all lines. Replace zeros with ones.
            // See also: https://msdn.microsoft.com/en-us/library/aa443952(v=vs.60).aspx

            var startLine = context.start.Line == 0 ? 1 : context.start.Line;
            var startCol = context.start.StartIndex == 0 ? 1 : context.start.StartIndex;
            var endLine = context.stop.Line == 0 ? 1 : context.stop.Line;
            var endCol = context.stop.StopIndex == 0 ? 1 : context.stop.StopIndex + 1;

            return new Selection(
                startLine,
                startCol, //todo: figure out why start col is off; see also https://github.com/retailcoder/Rubberduck/commit/e831a7aa1ced1498374f92a8dd3e9f37a587a3b8#commitcomment-9637108
                endLine,
                endCol
                );
        }

        /// <summary>
        /// Returns <c>true</c> if specified <c>Selection</c> contains this node.
        /// </summary>
        public static bool IsInSelection(this ParserRuleContext context, Selection selection)
        {
            var contextSelection = context.GetSelection();
            return selection.Contains(contextSelection);
        }

        public static VBAccessibility GetAccessibility(this VBParser.VisibilityContext context)
        {
            if (context == null)
                return VBAccessibility.Implicit;

            return (VBAccessibility) Enum.Parse(typeof (VBAccessibility), context.GetText());
        }

        public static string Signature(this VBParser.FunctionStmtContext context)
        {
            var visibility = context.Visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.AmbiguousIdentifier().GetText();
            var argsText = context.ArgList().GetText();
            
            var asType = context.AsTypeClause();
            var asTypeText = asType == null ? string.Empty : asType.GetText();

            return (visibilityText + ' ' + Tokens.Function + ' ' + identifierText + argsText + ' ' + asTypeText).Trim();
        }

        public static string Signature(this VBParser.SubStmtContext context)
        {
            var visibility = context.Visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.AmbiguousIdentifier().GetText();
            var argsText = context.ArgList().GetText();

            return (visibilityText + ' ' + Tokens.Sub + ' ' + identifierText + argsText).Trim();
        }

        public static string Signature(this VBParser.PropertyGetStmtContext context)
        {
            var visibility = context.Visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.AmbiguousIdentifier().GetText();
            var argsText = context.ArgList().GetText();

            var asType = context.AsTypeClause();
            var asTypeText = asType == null ? string.Empty : asType.GetText();

            return (visibilityText + ' ' + Tokens.Property + ' ' + Tokens.Get + ' ' + identifierText + argsText + ' ' + asTypeText).Trim();
        }

        public static string Signature(this VBParser.PropertyLetStmtContext context)
        {
            var visibility = context.Visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.AmbiguousIdentifier().GetText();
            var argsText = context.ArgList().GetText();

            return (visibilityText + ' ' + Tokens.Property + ' ' + Tokens.Let + ' ' + identifierText + argsText).Trim();
        }

        public static string Signature(this VBParser.PropertySetStmtContext context)
        {
            var visibility = context.Visibility();
            var visibilityText = visibility == null ? string.Empty : visibility.GetText();

            var identifierText = context.AmbiguousIdentifier().GetText();
            var argsText = context.ArgList().GetText();

            return (visibilityText + ' ' + Tokens.Property + ' ' + Tokens.Set + ' ' + identifierText + argsText).Trim();
        }
    }
}