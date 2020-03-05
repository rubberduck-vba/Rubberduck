using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Experimentals;
using Rubberduck.Parsing;

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
    internal sealed class EmptyWhileWendBlockInspection : ParseTreeInspectionBase<VBAParser.WhileWendStmtContext>
    {
        public EmptyWhileWendBlockInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new EmptyWhileWendBlockListener();
        }

        protected override IInspectionListener<VBAParser.WhileWendStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.WhileWendStmtContext> context)
        {
            return InspectionResults.EmptyWhileWendBlockInspection;
        }

        public class EmptyWhileWendBlockListener : EmptyBlockInspectionListenerBase<VBAParser.WhileWendStmtContext>
        {
            public override void EnterWhileWendStmt([NotNull] VBAParser.WhileWendStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
