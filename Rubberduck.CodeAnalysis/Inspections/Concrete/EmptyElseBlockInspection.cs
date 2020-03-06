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
    /// Identifies empty 'Else' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Empty code blocks are redundant, dead code that should be removed. They can also be misleading about their intent:
    /// an empty block may be signalling an unfinished thought or an oversight.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If foo Then
    ///         ' ...
    ///     Else
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If foo Then
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal sealed class EmptyElseBlockInspection : ParseTreeInspectionBase<VBAParser.ElseBlockContext>
    {
        public EmptyElseBlockInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new EmptyElseBlockListener();
        }

        protected override IInspectionListener<VBAParser.ElseBlockContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.ElseBlockContext> context)
        {
            return InspectionResults.EmptyElseBlockInspection;
        }

        private class EmptyElseBlockListener : EmptyBlockInspectionListenerBase<VBAParser.ElseBlockContext>
        {
            public override void EnterElseBlock([NotNull] VBAParser.ElseBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
