using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Common;
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
    /// Identifies empty 'Case' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Case blocks in VBA do not "fall through"; an empty 'Case' block might be hiding a bug.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Select Case foo
    ///         Case 0 ' empty block
    ///         Case Is > 0
    ///             Debug.Print foo ' does not run if foo is 0.
    ///     End Select
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Select Case foo
    ///         Case 0
    ///             '...code...
    ///         Case Is > 0
    ///             '...code...
    ///     End Select
    /// End Sub
    /// ]]>
    /// </example>
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal class EmptyCaseBlockInspection : ParseTreeInspectionBase
    {
        public EmptyCaseBlockInspection(RubberduckParserState state)
            : base(state) { }

        public override IInspectionListener Listener { get; } =
            new EmptyCaseBlockListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionResults.EmptyCaseBlockInspection,
                                                        result));
        }

        public class EmptyCaseBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterCaseClause([NotNull] VBAParser.CaseClauseContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
