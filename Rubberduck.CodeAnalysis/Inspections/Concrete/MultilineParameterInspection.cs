using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags parameters declared across multiple physical lines of code.
    /// </summary>
    /// <why>
    /// When splitting a long list of parameters across multiple lines, care should be taken to avoid splitting a parameter declaration in two.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long, ByVal _ 
    ///                              bar As Long)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long, _ 
    ///                        ByVal bar As Long)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class MultilineParameterInspection : ParseTreeInspectionBase
    {
        public MultilineParameterInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new ParameterListener();
        }
        
        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(context => new QualifiedContextInspectionResult(this,
                                                  string.Format(context.Context.GetSelection().LineCount > 3
                                                        ? RubberduckUI.EasterEgg_Continuator
                                                        : Resources.Inspections.InspectionResults.MultilineParameterInspection, ((VBAParser.ArgContext)context.Context).unrestrictedIdentifier().GetText()),
                                                  context));
        }

        public class ParameterListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts
                = new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }
            
            public override void ExitArg([NotNull] VBAParser.ArgContext context)
            {
                if (context.Start.Line != context.Stop.Line)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
