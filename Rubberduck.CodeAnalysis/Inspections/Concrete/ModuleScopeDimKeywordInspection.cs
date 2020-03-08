using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about module-level declarations made using the 'Dim' keyword.
    /// </summary>
    /// <why>
    /// Private module variables should be declared using the 'Private' keyword. While 'Dim' is also legal, it should preferably be 
    /// restricted to declarations of procedure-scoped local variables, for consistency, since public module variables are declared with the 'Public' keyword.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Dim foo As Long
    /// ' ...
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Private foo As Long
    /// ' ...
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ModuleScopeDimKeywordInspection : ParseTreeInspectionBase<VBAParser.VariableSubStmtContext>
    {
        public ModuleScopeDimKeywordInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new ModuleScopedDimListener();
        }

        protected override IInspectionListener<VBAParser.VariableSubStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.VariableSubStmtContext> context)
        {
            var identifierName = context.Context.identifier().GetText();
            return string.Format(
                InspectionResults.ModuleScopeDimKeywordInspection,
                identifierName);
        }

        private class ModuleScopedDimListener : InspectionListenerBase<VBAParser.VariableSubStmtContext>
        {
            public override void ExitVariableStmt([NotNull] VBAParser.VariableStmtContext context)
            {
                if (context.DIM() != null && context.TryGetAncestor<VBAParser.ModuleDeclarationsElementContext>(out _))
                {
                    var resultContexts = context.GetDescendents<VBAParser.VariableSubStmtContext>();
                    foreach (var resultContext in resultContexts)
                    {
                        SaveContext(resultContext);
                    }
                }
            }
        }
    }
}