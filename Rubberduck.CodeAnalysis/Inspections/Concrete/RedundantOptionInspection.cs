using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies redundant module options that are set to their implicit default.
    /// </summary>
    /// <why>
    /// Module options that are redundant can be safely removed. Disable this inspection if your convention is to explicitly specify them; a future 
    /// inspection may be used to enforce consistently explicit module options.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// Option Base 0
    /// Option Compare Binary
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class RedundantOptionInspection : ParseTreeInspectionBase
    {
        public RedundantOptionInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Listener = new RedundantModuleOptionListener();
        }

        public override IInspectionListener Listener { get; }
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return string.Format(
                InspectionResults.RedundantOptionInspection, 
                context.Context.GetText());
        }

        public class RedundantModuleOptionListener : InspectionListenerBase
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
