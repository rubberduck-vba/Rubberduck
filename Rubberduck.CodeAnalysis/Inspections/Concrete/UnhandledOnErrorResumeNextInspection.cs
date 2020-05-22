using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Finds instances of 'On Error Resume Next' that don't have a corresponding 'On Error GoTo 0' to restore error handling.
    /// </summary>
    /// <why>
    /// 'On Error Resume Next' should be constrained to a limited number of instructions, otherwise it supresses error handling 
    /// for the rest of the procedure; 'On Error GoTo 0' reinstates error handling. 
    /// This inspection helps treating 'Resume Next' and 'GoTo 0' as a code block (similar to 'With...End With'), essentially.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Error Resume Next ' error handling is never restored in this scope.
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Error Resume Next
    ///     ' ...
    ///     On Error GoTo 0
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UnhandledOnErrorResumeNextInspection : ParseTreeInspectionBase<VBAParser.OnErrorStmtContext, IReadOnlyList<VBAParser.OnErrorStmtContext>>
    {
        private readonly OnErrorStatementListener _listener = new OnErrorStatementListener();

        public UnhandledOnErrorResumeNextInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override IInspectionListener<VBAParser.OnErrorStmtContext> ContextListener => _listener;

        protected override string ResultDescription(QualifiedContext<VBAParser.OnErrorStmtContext> context, IReadOnlyList<VBAParser.OnErrorStmtContext> properties)
        {
            return InspectionResults.UnhandledOnErrorResumeNextInspection;
        }

        protected override (bool isResult, IReadOnlyList<VBAParser.OnErrorStmtContext> properties) IsResultContextWithAdditionalProperties(QualifiedContext<VBAParser.OnErrorStmtContext> context, DeclarationFinder finder)
        {
            return (true, _listener.UnhandledContexts(context));
        }

        private class OnErrorStatementListener : InspectionListenerBase<VBAParser.OnErrorStmtContext>
        {
            private readonly List<QualifiedContext<VBAParser.OnErrorStmtContext>> _unhandledContexts = new List<QualifiedContext<VBAParser.OnErrorStmtContext>>();
            private readonly Dictionary<QualifiedContext<VBAParser.OnErrorStmtContext>, List<VBAParser.OnErrorStmtContext>> _unhandledContextsMap = new Dictionary<QualifiedContext<VBAParser.OnErrorStmtContext>, List<VBAParser.OnErrorStmtContext>>();

            public IReadOnlyList<VBAParser.OnErrorStmtContext> UnhandledContexts(QualifiedContext<VBAParser.OnErrorStmtContext> context)
            {
                return _unhandledContextsMap.TryGetValue(context, out var unhandledContexts)
                    ? unhandledContexts
                    : new List<VBAParser.OnErrorStmtContext>();
            }

            public override void ClearContexts()
            {
                _unhandledContextsMap.Clear();
                base.ClearContexts();
            }

            public override void ClearContexts(QualifiedModuleName module)
            {
                var keysInModule = _unhandledContextsMap.Keys
                    .Where(context => context.ModuleName.Equals(module));

                foreach (var key in keysInModule)
                {
                    _unhandledContextsMap.Remove(key);
                }

                base.ClearContexts(module);
            }

            public override void ExitModuleBodyElement(VBAParser.ModuleBodyElementContext context)
            {
                if (_unhandledContexts.Any())
                {
                    foreach (var errorContext in _unhandledContexts)
                    {
                        _unhandledContextsMap.Add(errorContext, new List<VBAParser.OnErrorStmtContext>(_unhandledContexts.Select(ctx => ctx.Context)));
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

            private void SaveUnhandledContext(VBAParser.OnErrorStmtContext context)
            {
                var module = CurrentModuleName;
                var qualifiedContext = new QualifiedContext<VBAParser.OnErrorStmtContext>(module, context);
                _unhandledContexts.Add(qualifiedContext);
            }
        }
    }
}
