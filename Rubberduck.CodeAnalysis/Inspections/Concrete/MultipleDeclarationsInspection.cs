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
    /// Flags declaration statements spanning multiple physical lines of code.
    /// </summary>
    /// <why>
    /// Declaration statements should generally declare a single variable.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Dim foo As Long, bar As Long
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Dim foo As Long 
    /// Dim bar As Long 
    /// ]]>
    /// </example>
    public sealed class MultipleDeclarationsInspection : ParseTreeInspectionBase
    {
        public MultipleDeclarationsInspection(RubberduckParserState state)
            : base(state) { }

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.MultipleDeclarationsInspection;
        }

        public override IInspectionListener Listener { get; } = new ParameterListListener();

        public class ParameterListListener : InspectionListenerBase
        {
            public override void ExitVariableListStmt([NotNull] VBAParser.VariableListStmtContext context)
            {
                if (context.variableSubStmt().Length > 1)
                {
                   SaveContext(context);
                }
            }

            public override void ExitConstStmt([NotNull] VBAParser.ConstStmtContext context)
            {
                if (context.constSubStmt().Length > 1)
                {
                    SaveContext(context);
                }
            }
        }
    }
}
