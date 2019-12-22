using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Resources.Experimentals;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty 'While...Wend' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Dead code should be removed. A loop without a body is usually redundant.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     While foo < 100
    ///         'no executable statements... would be an infinite loop if entered
    ///     Wend
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     While foo < 100
    ///         foo = foo + 1
    ///     Wend
    /// End Sub
    /// ]]>
    /// </example>
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal class EmptyWhileWendBlockInspection : ParseTreeInspectionBase
    {
        public EmptyWhileWendBlockInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionResults.EmptyWhileWendBlockInspection,
                                                        result));
        }

        public override IInspectionListener Listener { get; } =
            new EmptyWhileWendBlockListener();

        public class EmptyWhileWendBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterWhileWendStmt([NotNull] VBAParser.WhileWendStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
