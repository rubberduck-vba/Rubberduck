using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies redundant module options that are set to their implicit default.
    /// </summary>
    /// <why>
    /// Module options that are redundant can be safely removed. Disable this inspection if your convention is to explicitly specify them; a future 
    /// inspection may be used to enforce consistently explicit module options.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Option Base 0
    /// Option Compare Binary
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class RedundantOptionInspection : ParseTreeInspectionBase<ParserRuleContext>
    {
        public RedundantOptionInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new RedundantModuleOptionListener();
        }

        protected override IInspectionListener<ParserRuleContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return string.Format(
                InspectionResults.RedundantOptionInspection, 
                context.Context.GetText());
        }

        private class RedundantModuleOptionListener : InspectionListenerBase<ParserRuleContext>
        {
            public override void ExitOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
            {
                if (context.numberLiteral()?.INTEGERLITERAL().Symbol.Text == "0")
                {
                   SaveContext(context);
                }
            }

            public override void ExitOptionCompareStmt(VBAParser.OptionCompareStmtContext context)
            {
                // BINARY is the default, and DATABASE is specified by default + only valid in Access.
                if (context.TEXT() == null && context.DATABASE() == null)
                {
                    SaveContext(context);
                }
            }
        }
    }
}
