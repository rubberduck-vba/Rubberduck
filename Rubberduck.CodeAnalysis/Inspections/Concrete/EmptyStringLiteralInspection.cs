using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags uses of an empty string literal ("").
    /// </summary>
    /// <why>
    /// Standard library constant 'vbNullString' is more explicit about its intent, and should be preferred to a string literal. 
    /// While the memory gain is meaningless, an empty string literal still takes up 2 bytes of memory,
    /// but 'vbNullString' is a null string pointer, and doesn't.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As String)
    ///     If foo = "" Then
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As String)
    ///     If foo = vbNullString Then
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class EmptyStringLiteralInspection : ParseTreeInspectionBase
    {
        public EmptyStringLiteralInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        public override IInspectionListener Listener { get; } =
            new EmptyStringLiteralListener();

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.EmptyStringLiteralInspection;
        }

        public class EmptyStringLiteralListener : InspectionListenerBase
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
