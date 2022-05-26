using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags parameters declared across multiple physical lines of code.
    /// </summary>
    /// <why>
    /// When splitting a long list of parameters across multiple lines, care should be taken to avoid splitting a parameter declaration in two.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long, ByVal _ 
    ///                              bar As Long)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long, _ 
    ///                        ByVal bar As Long)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class MultilineParameterInspection : ParseTreeInspectionBase<VBAParser.ArgContext>
    {
        public MultilineParameterInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new ParameterListener();
        }
        
        protected override IInspectionListener<VBAParser.ArgContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.ArgContext> context)
        {
            var parameterText = context.Context.unrestrictedIdentifier().GetText();
            return string.Format(
                context.Context.GetSelection().LineCount > 3
                    ? CodeAnalysisUI.EasterEgg_Continuator
                    : Resources.Inspections.InspectionResults.MultilineParameterInspection,
                parameterText);
        }

        private class ParameterListener : InspectionListenerBase<VBAParser.ArgContext>
        {
            public override void ExitArg([NotNull] VBAParser.ArgContext context)
            {
                if (context.Start.Line != context.Stop.Line)
                {
                    SaveContext(context);
                }
            }
        }
    }
}
