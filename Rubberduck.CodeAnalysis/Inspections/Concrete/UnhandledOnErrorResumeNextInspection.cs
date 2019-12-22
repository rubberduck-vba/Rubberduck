using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Finds instances of 'On Error Resume Next' that don't have a corresponding 'On Error GoTo 0' to restore error handling.
    /// </summary>
    /// <why>
    /// 'On Error Resume Next' should be constrained to a limited number of instructions, otherwise it supresses error handling 
    /// for the rest of the procedure; 'On Error GoTo 0' reinstates error handling. 
    /// This inspection helps treating 'Resume Next' and 'GoTo 0' as a code block (similar to 'With...End With'), essentially.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Error Resume Next ' error handling is never restored in this scope.
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Error Resume Next
    ///     ' ...
    ///     On Error GoTo 0
    /// End Sub
    /// ]]>
    /// </example>
    public class UnhandledOnErrorResumeNextInspection : ParseTreeInspectionBase
    {
        private readonly Dictionary<QualifiedContext<ParserRuleContext>, List<ParserRuleContext>> _unhandledContextsMap =
            new Dictionary<QualifiedContext<ParserRuleContext>, List<ParserRuleContext>>();

        public UnhandledOnErrorResumeNextInspection(RubberduckParserState state) : base(state)
        {
            Listener = new OnErrorStatementListener(_unhandledContextsMap);
        }

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(result =>
                {
                    dynamic properties = new PropertyBag();
                    properties.UnhandledContexts = _unhandledContextsMap[result];

                    return new QualifiedContextInspectionResult(this, InspectionResults.UnhandledOnErrorResumeNextInspection, result, properties);
                });
        }
    }

    public class OnErrorStatementListener : VBAParserBaseListener, IInspectionListener
    {
        private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
        private readonly List<QualifiedContext<ParserRuleContext>> _unhandledContexts = new List<QualifiedContext<ParserRuleContext>>();
        private readonly Dictionary<QualifiedContext<ParserRuleContext>, List<ParserRuleContext>> _unhandledContextsMap;

        public OnErrorStatementListener(Dictionary<QualifiedContext<ParserRuleContext>, List<ParserRuleContext>> unhandledContextsMap)
        {
            _unhandledContextsMap = unhandledContextsMap;
        }

        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

        public void ClearContexts()
        {
            _contexts.Clear();
            _unhandledContextsMap.Clear();
        }

        public QualifiedModuleName CurrentModuleName { get; set; }

        public override void ExitModuleBodyElement(VBAParser.ModuleBodyElementContext context)
        {
            if (_unhandledContexts.Any())
            {
                foreach (var errorContext in _unhandledContexts)
                {
                    _unhandledContextsMap.Add(errorContext, new List<ParserRuleContext>(_unhandledContexts.Select(ctx => ctx.Context)));
                }

                _contexts.AddRange(_unhandledContexts);

                _unhandledContexts.Clear();
            }
        }

        public override void ExitOnErrorStmt(VBAParser.OnErrorStmtContext context)
        {
            if (context.RESUME() != null)
            {
                _unhandledContexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
            else if (context.GOTO() != null)
            {
                _unhandledContexts.Clear();
            }
        }
    }
}
