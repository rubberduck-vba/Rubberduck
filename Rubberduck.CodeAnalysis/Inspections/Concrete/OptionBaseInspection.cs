using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags modules that specify Option Base 1.
    /// </summary>
    /// <why>
    /// Implicit array lower bound is 0 by default, and Option Base 1 makes it 1. While compelling in a 1-based environment like the Excel object model, 
    /// having an implicit lower bound of 1 for implicitly-sized user arrays does not change the fact that arrays are always better off with explicit boundaries.
    /// Because 0 is always the lower array bound in many other programming languages, this option may trip a reader/maintainer with a different background.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// Option Base 1
    /// 
    /// Public Sub DoSomething()
    ///     Dim foo(10) As Long ' implicit lower bound is 1, array has 10 items.
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// Public Sub DoSomething()
    ///     Dim foo(10) As Long ' implicit lower bound is 0, array has 11 items.
    ///     Dim bar(1 To 10) As Long ' explicit lower bound removes all ambiguities, Option Base is redundant.
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class OptionBaseInspection : ParseTreeInspectionBase
    {
        public OptionBaseInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new OptionBaseStatementListener();
        }
        
        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(context => new QualifiedContextInspectionResult(this,
                                                        string.Format(InspectionResults.OptionBaseInspection, context.ModuleName.ComponentName),
                                                        context));
        }

        public class OptionBaseStatementListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
            {
                if (context.numberLiteral()?.INTEGERLITERAL().Symbol.Text == "1")
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
