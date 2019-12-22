using Rubberduck.Inspections.Abstract;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Results;
using static Rubberduck.Parsing.Grammar.VBAParser;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates 'For' loops where the 'Step' token is omitted.
    /// </summary>
    /// <why>
    /// Out of convention or preference, explicit 'Step' specifiers could be considered mandatory; 
    /// this inspection can ensure the consistency of the convention.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 100 ' Step is implicitly 1
    ///         ' ...
    ///     Next
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 100 Step 1 ' explicit 'Step 1' could also be considered redundant.
    ///         ' ...
    ///     Next
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class StepIsNotSpecifiedInspection : ParseTreeInspectionBase
    {
        public StepIsNotSpecifiedInspection(RubberduckParserState state) : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionResults.StepIsNotSpecifiedInspection,
                                                        result));
        }

        public override IInspectionListener Listener { get; } =
            new StepIsNotSpecifiedListener();
    }

    public class StepIsNotSpecifiedListener : VBAParserBaseListener, IInspectionListener
    {
        private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

        public QualifiedModuleName CurrentModuleName
        {
            get;
            set;
        }

        public void ClearContexts()
        {
            _contexts.Clear();
        }

        public override void EnterForNextStmt([NotNull] ForNextStmtContext context)
        {
            StepStmtContext stepStatement = context.stepStmt();

            if (stepStatement == null)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
    }
}
