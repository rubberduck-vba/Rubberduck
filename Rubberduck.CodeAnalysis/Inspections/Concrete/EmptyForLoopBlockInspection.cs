using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty 'For...Next' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Dead code should be removed. A loop without a body is usually redundant.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Dim i As Long
    ///     For i = 0 To foo
    ///         ' no executable statement...
    ///     Next
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Dim i As Long
    ///     For i = 0 To foo
    ///         Debug.Print i
    ///     Next
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class EmptyForLoopBlockInspection : EmptyBlockInspectionBase<VBAParser.ForNextStmtContext>
    {
        public EmptyForLoopBlockInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new EmptyForLoopBlockListener();
        }

        protected override IInspectionListener<VBAParser.ForNextStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.ForNextStmtContext> context)
        {
            return InspectionResults.EmptyForLoopBlockInspection;
        }

        private class EmptyForLoopBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context)
            {
                InspectBlockForExecutableStatements(context.unterminatedBlock(), context);
            }
        }
    }
}
