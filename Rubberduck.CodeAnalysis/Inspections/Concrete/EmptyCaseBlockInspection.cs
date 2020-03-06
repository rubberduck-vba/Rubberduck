using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Experimentals;
using Rubberduck.Parsing;

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
    internal sealed class EmptyCaseBlockInspection : ParseTreeInspectionBase<VBAParser.CaseClauseContext>
    {
        public EmptyCaseBlockInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new EmptyCaseBlockListener();
        }

        protected override IInspectionListener<VBAParser.CaseClauseContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.CaseClauseContext> context)
        {
            return InspectionResults.EmptyCaseBlockInspection;
        }

        private class EmptyCaseBlockListener : EmptyBlockInspectionListenerBase<VBAParser.CaseClauseContext>
        {
            public override void EnterCaseClause([NotNull] VBAParser.CaseClauseContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
