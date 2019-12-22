using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Resources.Experimentals;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty 'If' blocks.
    /// </summary>
    /// <why>
    /// Conditional expression is inverted; there would not be a need for an 'Else' block otherwise.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If foo Then
    ///     Else
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If Not foo Then
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal class EmptyIfBlockInspection : ParseTreeInspectionBase
    {
        public EmptyIfBlockInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionResults.EmptyIfBlockInspection,
                                                        result));
        }

        public override IInspectionListener Listener { get; } =
            new EmptyIfBlockListener();

        public class EmptyIfBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterIfStmt([NotNull] VBAParser.IfStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }

            public override void EnterElseIfBlock([NotNull] VBAParser.ElseIfBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }

            public override void EnterSingleLineIfStmt([NotNull] VBAParser.SingleLineIfStmtContext context)
            {
                if (context.ifWithEmptyThen() != null)
                {
                    AddResult(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context.ifWithEmptyThen()));
                }
            }
        }
    }
}
