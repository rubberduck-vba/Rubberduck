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
    /// Identifies empty 'Do...Loop While' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Dead code should be removed. A loop without a body is usually redundant.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Do
    ///         ' no executable statement...
    ///     Loop While foo < 100
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Do
    ///         Debug.Print foo
    ///     Loop While foo < 100
    /// End Sub
    /// ]]>
    /// </example>
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal class EmptyDoWhileBlockInspection : ParseTreeInspectionBase
    {
        public EmptyDoWhileBlockInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(result => 
                    new QualifiedContextInspectionResult(this, InspectionResults.EmptyDoWhileBlockInspection, result));
        }

        public override IInspectionListener Listener { get; } = new EmptyDoWhileBlockListener();

        public class EmptyDoWhileBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterDoLoopStmt([NotNull] VBAParser.DoLoopStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
