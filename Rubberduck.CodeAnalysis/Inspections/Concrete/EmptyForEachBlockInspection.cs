using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty 'For Each...Next' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Dead code should be removed. A loop without a body is usually redundant.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim sheet As Worksheet
    ///     For Each sheet In ThisWorkbook.Worksheets
    ///         ' no executable statement...
    ///     Next
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim sheet As Worksheet
    ///     For Each sheet In ThisWorkbook.Worksheets
    ///         Debug.Print sheet.Name
    ///     Next
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class EmptyForEachBlockInspection : EmptyBlockInspectionBase<VBAParser.ForEachStmtContext>
    {
        public EmptyForEachBlockInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new EmptyForEachBlockListener();
        }

        protected override IInspectionListener<VBAParser.ForEachStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.ForEachStmtContext> context)
        {
            return InspectionResults.EmptyForEachBlockInspection;
        }

        private class EmptyForEachBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterForEachStmt([NotNull] VBAParser.ForEachStmtContext context)
            {
                InspectBlockForExecutableStatements(context.unterminatedBlock(), context);
            }
        }
    }
}
