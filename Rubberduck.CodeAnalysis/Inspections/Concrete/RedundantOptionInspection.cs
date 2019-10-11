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
        public RedundantOptionInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new RedundantModuleOptionListener();
        }

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                                   .Select(context => new QualifiedContextInspectionResult(this,
                                                                           string.Format(InspectionResults.RedundantOptionInspection, context.Context.GetText()),
                                                                           context));
        }

        public class RedundantModuleOptionListener : VBAParserBaseListener, IInspectionListener
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
                if (context.numberLiteral()?.INTEGERLITERAL().Symbol.Text == "0")
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }

            public override void ExitOptionCompareStmt(VBAParser.OptionCompareStmtContext context)
            {
                // BINARY is the default, and DATABASE is specified by default + only valid in Access.
                if (context.TEXT() == null && context.DATABASE() == null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
