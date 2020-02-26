using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about module-level declarations made using the 'Dim' keyword.
    /// </summary>
    /// <why>
    /// Private module variables should be declared using the 'Private' keyword. While 'Dim' is also legal, it should preferably be 
    /// restricted to declarations of procedure-scoped local variables, for consistency, since public module variables are declared with the 'Public' keyword.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// Dim foo As Long
    /// ' ...
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// Private foo As Long
    /// ' ...
    /// ]]>
    /// </example>
    public sealed class ModuleScopeDimKeywordInspection : ParseTreeInspectionBase
    {
        public ModuleScopeDimKeywordInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        public override IInspectionListener Listener { get; } = new ModuleScopedDimListener();
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            var identifierName = ((VBAParser.VariableSubStmtContext) context.Context).identifier().GetText();
            return string.Format(
                InspectionResults.ModuleScopeDimKeywordInspection,
                identifierName);
        }

        public class ModuleScopedDimListener : InspectionListenerBase
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