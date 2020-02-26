using System.Collections.Generic;
using System.Linq;
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
    public class UnhandledOnErrorResumeNextInspection : ParseTreeInspectionBase<IReadOnlyList<ParserRuleContext>>
    {
        private readonly Dictionary<QualifiedContext<ParserRuleContext>, List<ParserRuleContext>> _unhandledContextsMap =
            new Dictionary<QualifiedContext<ParserRuleContext>, List<ParserRuleContext>>();

        public UnhandledOnErrorResumeNextInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Listener = new OnErrorStatementListener(_unhandledContextsMap);
        }

        public override IInspectionListener Listener { get; }
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context, IReadOnlyList<ParserRuleContext> properties)
        {
            return InspectionResults.UnhandledOnErrorResumeNextInspection;
        }

        protected override (bool isResult, IReadOnlyList<ParserRuleContext> properties) IsResultContextWithAdditionalProperties(QualifiedContext<ParserRuleContext> context)
        {
            return (true, _unhandledContextsMap[context]);
        }
    }

    public class OnErrorStatementListener : InspectionListenerBase
    {
        private readonly List<QualifiedContext<ParserRuleContext>> _unhandledContexts = new List<QualifiedContext<ParserRuleContext>>();
        private readonly Dictionary<QualifiedContext<ParserRuleContext>, List<ParserRuleContext>> _unhandledContextsMap;

        public OnErrorStatementListener(Dictionary<QualifiedContext<ParserRuleContext>, List<ParserRuleContext>> unhandledContextsMap)
        {
            _unhandledContextsMap = unhandledContextsMap;
        }

        public override void ClearContexts()
        {
            _unhandledContextsMap.Clear();
            base.ClearContexts();
        }

        public override void ExitModuleBodyElement(VBAParser.ModuleBodyElementContext context)
        {
            if (_unhandledContexts.Any())
            {
                foreach (var errorContext in _unhandledContexts)
                {
                    _unhandledContextsMap.Add(errorContext, new List<ParserRuleContext>(_unhandledContexts.Select(ctx => ctx.Context)));
                    SaveContext(errorContext.Context);
                }

                _unhandledContexts.Clear();
            }
        }

        public override void ExitOnErrorStmt(VBAParser.OnErrorStmtContext context)
        {
            if (context.RESUME() != null)
            {
                SaveUnhandledContext(context);
            }
            else if (context.GOTO() != null)
            {
                _unhandledContexts.Clear();
            }
        }

        private void SaveUnhandledContext(ParserRuleContext context)
        {
            var module = CurrentModuleName;
            var qualifiedContext = new QualifiedContext<ParserRuleContext>(module, context);
            _unhandledContexts.Add(qualifiedContext);
        }
    }
}
