using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty 'Do...Loop While' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Dead code should be removed. A loop without a body is usually redundant.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Do
    ///         ' no executable statement...
    ///     Loop While foo < 100
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Do
    ///         Debug.Print foo
    ///     Loop While foo < 100
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class EmptyDoWhileBlockInspection : EmptyBlockInspectionBase<VBAParser.DoLoopStmtContext>
    {
        public EmptyDoWhileBlockInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new EmptyDoWhileBlockListener();
        }

        protected override IInspectionListener<VBAParser.DoLoopStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.DoLoopStmtContext> context)
        {
            return InspectionResults.EmptyDoWhileBlockInspection;
        }

        private class EmptyDoWhileBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterDoLoopStmt([NotNull] VBAParser.DoLoopStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
