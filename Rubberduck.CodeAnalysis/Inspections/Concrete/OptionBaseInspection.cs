using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags modules that specify Option Base 1.
    /// </summary>
    /// <why>
    /// Implicit array lower bound is 0 by default, and Option Base 1 makes it 1. While compelling in a 1-based environment like the Excel object model, 
    /// having an implicit lower bound of 1 for implicitly-sized user arrays does not change the fact that arrays are always better off with explicit boundaries.
    /// Because 0 is always the lower array bound in many other programming languages, this option may trip a reader/maintainer with a different background.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Option Base 1
    /// 
    /// Public Sub DoSomething()
    ///     Dim foo(10) As Long ' implicit lower bound is 1, array has 10 items.
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Public Sub DoSomething()
    ///     Dim foo(10) As Long ' implicit lower bound is 0, array has 11 items.
    ///     Dim bar(1 To 10) As Long ' explicit lower bound removes all ambiguities, Option Base is redundant.
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class OptionBaseInspection : ParseTreeInspectionBase<VBAParser.OptionBaseStmtContext>
    {
        public OptionBaseInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new OptionBaseStatementListener();
        }
        
        protected override IInspectionListener<VBAParser.OptionBaseStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.OptionBaseStmtContext> context)
        {
            var moduleName = context.ModuleName.ComponentName;
            return string.Format(
                InspectionResults.OptionBaseInspection, 
                moduleName);
        }

        private class OptionBaseStatementListener : InspectionListenerBase<VBAParser.OptionBaseStmtContext>
        {
            public override void ExitOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
            {
                if (context.numberLiteral()?.INTEGERLITERAL().Symbol.Text == "1")
                {
                   SaveContext(context);
                }
            }
        }
    }
}
