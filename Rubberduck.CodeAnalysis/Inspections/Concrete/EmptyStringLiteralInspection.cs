using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

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
        public EmptyStringLiteralInspection(RubberduckParserState state)
            : base(state) { }

        public override IInspectionListener Listener { get; } =
            new EmptyStringLiteralListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       InspectionResults.EmptyStringLiteralInspection,
                                                       result));
        }

        public class EmptyStringLiteralListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;
            
            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitLiteralExpression(VBAParser.LiteralExpressionContext context)
            {
                var literal = context.STRINGLITERAL();
                if (literal != null && literal.GetText() == "\"\"")
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
