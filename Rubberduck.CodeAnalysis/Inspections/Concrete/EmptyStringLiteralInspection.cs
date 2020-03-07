using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags uses of an empty string literal ("").
    /// </summary>
    /// <why>
    /// Standard library constant 'vbNullString' is more explicit about its intent, and should be preferred to a string literal. 
    /// While the memory gain is meaningless, an empty string literal still takes up 2 bytes of memory,
    /// but 'vbNullString' is a null string pointer, and doesn't.
    /// </why>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As String)
    ///     If foo = "" Then
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As String)
    ///     If foo = vbNullString Then
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    internal sealed class EmptyStringLiteralInspection : ParseTreeInspectionBase<VBAParser.LiteralExpressionContext>
    {
        public EmptyStringLiteralInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new EmptyStringLiteralListener();
        }

        protected override IInspectionListener<VBAParser.LiteralExpressionContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.LiteralExpressionContext> context)
        {
            return InspectionResults.EmptyStringLiteralInspection;
        }

        private class EmptyStringLiteralListener : InspectionListenerBase<VBAParser.LiteralExpressionContext>
        {
            public override void ExitLiteralExpression(VBAParser.LiteralExpressionContext context)
            {
                var literal = context.STRINGLITERAL();
                if (literal != null && literal.GetText() == "\"\"")
                {
                    SaveContext(context);
                }
            }
        }
    }
}
