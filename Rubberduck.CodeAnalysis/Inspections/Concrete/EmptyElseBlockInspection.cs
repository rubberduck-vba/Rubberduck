using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty 'Else' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Empty code blocks are redundant, dead code that should be removed. They can also be misleading about their intent:
    /// an empty block may be signalling an unfinished thought or an oversight.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If foo Then
    ///         ' ...
    ///     Else
    ///     End If
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If foo Then
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class EmptyElseBlockInspection : EmptyBlockInspectionBase<VBAParser.ElseBlockContext>
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

        private class EmptyElseBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterElseBlock([NotNull] VBAParser.ElseBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
