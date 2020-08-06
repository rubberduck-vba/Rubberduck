using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty 'Case' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Case blocks in VBA do not "fall through"; an empty 'Case' block might be hiding a bug.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Select Case foo
    ///         Case 0 ' empty block
    ///         Case Is > 0
    ///             Debug.Print foo ' does not run if foo is 0.
    ///     End Select
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
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
    /// </module>
    /// </example>
    internal sealed class EmptyCaseBlockInspection : EmptyBlockInspectionBase<VBAParser.CaseClauseContext>
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

        private class EmptyCaseBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterCaseClause([NotNull] VBAParser.CaseClauseContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
