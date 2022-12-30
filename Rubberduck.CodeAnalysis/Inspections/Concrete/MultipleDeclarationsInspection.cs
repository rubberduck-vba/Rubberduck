using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags declaration statements declaring multiple variables.
    /// </summary>
    /// <why>
    /// Declaration statements should generally declare a single variable. 
    /// Although this inspection does not take variable types into account, it is a common mistake to only declare an explicit type on the last variable in a list.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// 'note: RowNumber is untyped / implicitly Variant
    /// Dim RowNumber, ColumnNumber As Long
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Dim RowNumber As Long, ColumnNumber As Long
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Dim RowNumber As Long 
    /// Dim ColumnNumber As Long 
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class MultipleDeclarationsInspection : ParseTreeInspectionBase<ParserRuleContext>
    {
        public MultipleDeclarationsInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new ParameterListListener();
        }

        protected override IInspectionListener<ParserRuleContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.MultipleDeclarationsInspection;
        }

        private class ParameterListListener : InspectionListenerBase<ParserRuleContext>
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
